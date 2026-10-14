/* Budgie — a local-first monthly budget PWA.
 *
 * Data lives in localStorage on each device. With sync turned on, the whole
 * document is encrypted on-device (AES-GCM, key derived from the sync code)
 * and stored by a tiny Cloudflare Worker, so phone, laptop and family devices
 * converge. Bank SMS can be forwarded to the Worker by an iOS Shortcut or an
 * Android automation; the app pulls those, parses them and auto-categorises.
 */
(function () {
  "use strict";

  // Set this after deploying the Worker (see README). It can also be changed in Settings.
  var DEFAULT_SYNC_URL = "https://budgie-sync.benjasondixon.workers.dev";

  var STORE_KEY = "budgie.db.v1";
  var LOCAL_KEY = "budgie.local.v1";
  var SHEETJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  var TRANSFER = "__transfer";

  // ======================================================================
  // Utilities
  // ======================================================================
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var now = function () { return Date.now(); };
  function uid() { return Math.random().toString(36).slice(2, 9) + now().toString(36).slice(-5); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }
  function hash(str) {
    var h = 5381, s = String(str).toLowerCase().replace(/\s+/g, " ").trim();
    for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0;
    return h.toString(36);
  }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function iso(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function fromIso(s) { var p = s.split("-"); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function today() { var d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function dayDiff(a, b) { return Math.round((b - a) / 86400000); }
  var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var MON3 = MONTHS.map(function (m) { return m.slice(0, 3).toLowerCase(); });

  function money(n, cents) {
    var cur = (db && db.settings.currency) || "ZAR";
    try {
      return new Intl.NumberFormat("en-ZA", {
        style: "currency", currency: cur,
        minimumFractionDigits: cents ? 2 : 0, maximumFractionDigits: cents ? 2 : 0
      }).format(n || 0).replace(/ /g, " ");
    } catch (e) { return "R " + Math.round(n || 0); }
  }
  // Accepts "1 234,56", "1,234.56", "R1234", "-45.00", 1234
  function parseAmount(v) {
    if (typeof v === "number") return isFinite(v) ? v : NaN;
    var s = String(v == null ? "" : v).replace(/[R$€£\s ]|ZAR/gi, "");
    if (!s) return NaN;
    var neg = /^\(.*\)$/.test(s) || /^-/.test(s);
    s = s.replace(/[()\-+]/g, "");
    var lc = s.lastIndexOf(","), ld = s.lastIndexOf(".");
    if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    else if (lc > -1) s = /,\d{1,2}$/.test(s) ? s.replace(/,(?=\d{1,2}$)/, ".").replace(/,/g, "") : s.replace(/,/g, "");
    var n = parseFloat(s);
    return isNaN(n) ? NaN : (neg ? -n : n);
  }

  // ======================================================================
  // State
  // ======================================================================
  function blankDb() {
    return {
      v: 1,
      settings: { currency: "ZAR", startDay: 1, income: 0, updatedAt: 0 },
      categories: [], transactions: [], rules: []
    };
  }
  function load() {
    try { var raw = localStorage.getItem(STORE_KEY); if (raw) { var d = JSON.parse(raw); return Object.assign(blankDb(), d); } } catch (e) {}
    return blankDb();
  }
  function loadLocal() {
    var base = { theme: "auto", simple: false, syncUrl: DEFAULT_SYNC_URL, syncCode: "", lastSync: 0, onboarded: false };
    try { var raw = localStorage.getItem(LOCAL_KEY); if (raw) return Object.assign(base, JSON.parse(raw)); } catch (e) {}
    return base;
  }
  var db = load();
  var local = loadLocal();
  if (!local.syncUrl) local.syncUrl = DEFAULT_SYNC_URL;
  var ui = { tab: "home", offset: 0, actFilter: "all", actQuery: "", sheet: null };

  function persist() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(db)); } catch (e) { toast("Couldn't save on this device — storage is full or blocked."); }
  }
  function saveLocal() { try { localStorage.setItem(LOCAL_KEY, JSON.stringify(local)); } catch (e) {} }
  function commit() { persist(); scheduleSync(); render(); }
  function touch(o) { o.updatedAt = now(); return o; }

  function cats() { return db.categories.filter(function (c) { return !c.deleted; }).sort(function (a, b) { return (a.order || 0) - (b.order || 0); }); }
  function cat(id) { for (var i = 0; i < db.categories.length; i++) if (db.categories[i].id === id && !db.categories[i].deleted) return db.categories[i]; return null; }
  function txs() { return db.transactions.filter(function (t) { return !t.deleted; }); }
  function reviewQueue() { return txs().filter(function (t) { return t.needsReview; }).sort(function (a, b) { return a.date < b.date ? 1 : -1; }); }
  function groups() {
    var g = [];
    cats().forEach(function (c) { var n = c.group || "Expenses"; if (g.indexOf(n) < 0) g.push(n); });
    return g;
  }

  // ======================================================================
  // Budget periods (supports a payday start, e.g. the 25th)
  // ======================================================================
  function period(offset) {
    var sd = clamp(+db.settings.startDay || 1, 1, 28);
    var t = today();
    var start = t.getDate() >= sd ? new Date(t.getFullYear(), t.getMonth(), sd) : new Date(t.getFullYear(), t.getMonth() - 1, sd);
    start = new Date(start.getFullYear(), start.getMonth() + (offset || 0), sd);
    var end = new Date(start.getFullYear(), start.getMonth() + 1, sd - 1);
    if (sd === 1) end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
    var days = dayDiff(start, end) + 1;
    var elapsed = offset < 0 ? days : offset > 0 ? 0 : dayDiff(start, t) + 1;
    var named = sd > 15 ? end : start;
    var short = function (d) { return d.getDate() + " " + MONTHS[d.getMonth()].slice(0, 3); };
    return {
      start: start, end: end, s: iso(start), e: iso(end), days: days, elapsed: elapsed,
      left: days - elapsed, current: !offset, offset: offset || 0,
      label: MONTHS[named.getMonth()] + " " + named.getFullYear(),
      sub: sd === 1 ? "" : short(start) + " – " + short(end)
    };
  }
  function inPeriod(t, p) { return t.date >= p.s && t.date <= p.e; }
  function stats(p) {
    var list = txs().filter(function (t) { return inPeriod(t, p); });
    var by = {}, spent = 0, income = 0, uncategorised = 0;
    list.forEach(function (t) {
      if (t.type === "income") { income += t.amount; return; }
      if (t.type === "transfer" || t.catId === TRANSFER) return;
      spent += t.amount;
      if (t.catId && cat(t.catId)) by[t.catId] = (by[t.catId] || 0) + t.amount;
      else uncategorised += t.amount;
    });
    var budget = cats().reduce(function (s, c) { return s + (+c.budget || 0); }, 0);
    return { list: list, by: by, spent: spent, income: income, uncategorised: uncategorised, budget: budget };
  }
  // Fixed bills (rent, insurance, debit orders) leave in one go, so they
  // shouldn't make you look "ahead of pace". Only variable spending is paced.
  var FIXED_KINDS = ["housing", "insurance", "utilities", "phone", "subscriptions", "savings", "debt", "education"];
  function isFixed(c) {
    if (c.fixed === true || c.fixed === false) return c.fixed;
    if (/fixed|bill|debit order|commit|contract|obligation/i.test(c.group || "")) return true;
    return FIXED_KINDS.indexOf(c.kind || guessKind(c.name)) > -1;
  }
  function expectedByNow(c, spent, p) {
    var b = +c.budget || 0;
    if (!p.current) return p.offset < 0 ? b : 0;
    return isFixed(c) ? Math.min(spent, b) : b * p.elapsed / p.days;
  }
  // status of a category relative to time elapsed in the month
  function paceStatus(spent, budget, p, fixed) {
    if (!budget) return spent > 0 ? "over" : "ok";
    var used = spent / budget, time = p.elapsed / p.days;
    if (used > 1.0001) return "over";
    if (fixed) return "ok";
    if (p.current && used > time + 0.12 && used > 0.25) return "warn";
    if (!p.current && used > 0.9) return "warn";
    return "ok";
  }
  var STATUS_LABEL = { ok: "✓ On track", warn: "▲ Spending fast", over: "● Over budget" };

  // ======================================================================
  // Category knowledge (South African merchants) + auto-categoriser
  // ======================================================================
  var KINDS = {
    housing: { e: "🏠", syn: ["rent", "bond", "housing", "home loan", "levy", "levies", "mortgage", "accommodation"], m: ["rent", "bond", "home loan", "levy", "levies", "body corporate", "property"] },
    groceries: { e: "🛒", syn: ["grocer", "food", "supermarket", "household"], m: ["checkers", "shoprite", "pick n pay", "picknpay", "pnp", "woolworths", "woolies", "spar", "food lover", "usave", "boxer", "makro", "sixty60", "ok foods", "fruit & veg", "fruit and veg", "president hyper", "cambridge food", "superspar", "kwikspar", "food store", "butcher"] },
    eating: { e: "🍔", syn: ["eating", "restaurant", "takeaway", "take-away", "take away", "dining", "eat out", "coffee", "fast food", "meals out"], m: ["kfc", "mcdonald", "mcd ", "nando", "steers", "wimpy", "spur", "debonair", "roman", "ocean basket", "mugg", "vida e", "starbucks", "uber eats", "ubereats", "mr d", "mrd", "burger king", "pizza", "restaurant", "cafe", "coffee", "bakery", "tasha", "fishaways", "chicken licken", "galito", "rocomamas", "panarottis", "john dory", "seattle", "bootlegger", "kauai", "sushi", "grill", "bistro", "doppio", "col'cacchio", "mugg & bean", "mugg&bean"] },
    fuel: { e: "⛽", syn: ["fuel", "petrol", "diesel", "gas"], m: ["engen", "shell", "bp ", "sasol", "caltex", "totalenergies", "total ", "astron", "fuel", "garage", "puma energy", "petroport"] },
    transport: { e: "🚕", syn: ["transport", "travel", "uber", "taxi", "commute", "parking", "toll"], m: ["uber", "bolt", "gautrain", "taxi", "toll", "sanral", "parking", "e-toll", "myciti", "metrorail", "flysafair", "airlink", "lift airline", "kulula", "intercape", "greyhound"] },
    car: { e: "🚗", syn: ["car", "vehicle", "auto", "maintenance"], m: ["wesbank", "mfc", "vehicle finance", "tiger wheel", "supa quick", "midas", "autozone", "goldwagen", "car wash", "licence", "license renewal", "dekra"] },
    utilities: { e: "💡", syn: ["electric", "water", "utilit", "municipal", "rates", "power"], m: ["eskom", "city of", "municipality", "prepaid elec", "electricity", "city power", "water", "rates", "lightstone", "ekurhuleni", "tshwane", "joburg", "cape town"] },
    phone: { e: "📱", syn: ["phone", "cell", "internet", "wifi", "data", "airtime", "mobile", "fibre"], m: ["vodacom", "mtn", "cell c", "telkom", "rain", "afrihost", "webafrica", "airtime", "data bundle", "vumatel", "openserve", "coolideas", "mweb", "herotel", "frogfoot"] },
    insurance: { e: "🛡️", syn: ["insurance", "cover", "assurance", "funeral"], m: ["insurance", "outsurance", "santam", "old mutual", "sanlam", "liberty", "discovery life", "momentum", "king price", "miway", "hollard", "budget ins", "1life", "first for women", "dial direct", "auto & general", "clientele", "avbob", "assupol"] },
    medical: { e: "💊", syn: ["medical", "health", "pharmacy", "doctor", "medicine", "aid"], m: ["dis-chem", "dischem", "clicks", "pharmacy", "doctor", "dr ", "medical", "hospital", "netcare", "mediclinic", "life healthcare", "discovery health", "bonitas", "gems", "fedhealth", "bestmed", "medshield", "dentist", "optometrist", "spec-savers", "specsavers", "pathcare", "lancet"] },
    subscriptions: { e: "📺", syn: ["subscription", "streaming", "netflix", "apps", "software", "tv"], m: ["netflix", "showmax", "spotify", "apple.com", "itunes", "google", "youtube", "dstv", "multichoice", "disney", "prime video", "amazon prime", "icloud", "microsoft", "openai", "chatgpt", "anthropic", "claude.ai", "adobe", "deezer", "playstation plus", "xbox game pass", "canva"] },
    shopping: { e: "🛍️", syn: ["shopping", "clothes", "clothing", "personal", "fashion", "household items"], m: ["takealot", "superbalist", "mr price", "mrp", "ackermans", "pep ", "pepstores", "edgars", "foschini", "truworths", "h&m", "cotton on", "game ", "builders", "incredible connection", "amazon", "shein", "temu", "sportscene", "totalsports", "cna", "exclusive books", "markham", "jet ", "identity", "woolworths clothing", "zara", "bash", "makro", "hifi corp", "istore", "dion wired", "loot", "yuppiechef", "mr price home", "sheet street", "@home", "decofurn", "leroy merlin", "cashbuild", "mica"] },
    entertainment: { e: "🎉", syn: ["entertainment", "fun", "leisure", "hobbies", "going out", "social"], m: ["ster-kinekor", "ster kinekor", "nu metro", "computicket", "ticketpro", "quicket", "howler", "steam", "playstation", "xbox", "bowling", "tops", "liquor", "bottle store", "pub", "bar ", "tavern", "casino"] },
    personal: { e: "💇", syn: ["personal care", "beauty", "hair", "grooming", "toiletries"], m: ["salon", "barber", "hair", "nails", "spa", "beauty"] },
    fitness: { e: "🏋️", syn: ["gym", "fitness", "sport", "exercise"], m: ["virgin active", "planet fitness", "gym", "crossfit", "parkrun", "strava", "zwift", "sweat1000"] },
    education: { e: "🎓", syn: ["education", "school", "study", "fees", "tuition", "kids"], m: ["school", "university", "unisa", "varsity", "college", "tuition", "udemy", "coursera"] },
    savings: { e: "🏦", syn: ["saving", "invest", "emergency", "retirement", "tfsa", "goal"], m: ["savings", "invest", "easyequities", "easy equities", "satrix", "tfsa", "allan gray", "coronation", "10x", "sygnia", "etfsa", "retirement annuity"] },
    giving: { e: "🎁", syn: ["gift", "giving", "donation", "charity", "church", "tithe"], m: ["church", "tithe", "donation", "givengain", "backabuddy", "charity"] },
    pets: { e: "🐾", syn: ["pet", "dog", "cat", "vet"], m: ["vet", "pet", "petshop", "petworld", "absolute pets", "montana pet"] },
    fees: { e: "🏧", syn: ["bank fee", "fees", "charges", "bank charges", "interest"], m: ["fee", "service charge", "monthly acc", "admin charge", "interest charged", "sms notif", "cash dep fee"] },
    cash: { e: "💵", syn: ["cash", "atm", "withdrawal"], m: ["atm", "cash withdrawal", "cash wd", "withdrawal", "cash send"] },
    debt: { e: "💳", syn: ["debt", "loan", "credit card", "repayment", "store card"], m: ["loan", "credit card", "repayment", "store card", "personal loan", "edgars account", "truworths account"] }
  };
  function guessKind(name) {
    var n = " " + String(name || "").toLowerCase() + " ";
    var best = null, bestLen = 0;
    Object.keys(KINDS).forEach(function (k) {
      KINDS[k].syn.forEach(function (s) { if (n.indexOf(s) > -1 && s.length > bestLen) { best = k; bestLen = s.length; } });
    });
    return best;
  }
  function emojiFor(name) { var k = guessKind(name); return k ? KINDS[k].e : "🏷️"; }

  // Turns "POS PURCHASE CHECKERS HYPER 1234 SANDTON" into "checkers hyper"
  function merchantKey(m) {
    var w = String(m || "").toLowerCase()
      .replace(/\b(pos|purchase|pur|card|debit|credit|payment|pmt|ref|reference|online|int|za|zaf|rsa|pty|ltd|cc|inc|the)\b/g, " ")
      .replace(/[^a-z& ]+/g, " ").replace(/\s+/g, " ").trim().split(" ").filter(function (x) { return x.length > 1; });
    return w.slice(0, 2).join(" ");
  }

  function categorise(tx) {
    var res = { catId: null, confidence: 0, suggest: [], reason: "" };
    if (tx.type === "income") return { catId: null, confidence: 1, suggest: [], reason: "income" };
    var key = merchantKey(tx.merchant);
    var all = cats();
    // 1. a rule learned from you
    if (key) {
      var rule = db.rules.filter(function (r) { return r.id === key && !r.deleted; })[0];
      if (rule && (rule.catId === TRANSFER || cat(rule.catId))) {
        return { catId: rule.catId, confidence: 1, suggest: [rule.catId], reason: "You chose this before for “" + key + "”" };
      }
    }
    // same brand, different branch ("makro woodmead" vs "makro strubens") → strong suggestion only
    var first = key.split(" ")[0], near = null;
    if (first.length >= 4) near = db.rules.filter(function (r) { return !r.deleted && r.id.split(" ")[0] === first && cat(r.catId); })[0] || null;
    var hay = (" " + (tx.merchant || "") + " " + (tx.raw || "")).toLowerCase();
    var merch = (" " + (tx.merchant || "") + " ").toLowerCase();
    // 2. your own keywords on a category
    var kwHits = all.filter(function (c) {
      return String(c.keywords || "").split(",").map(function (s) { return s.trim().toLowerCase(); })
        .some(function (k) { return k.length > 1 && hay.indexOf(k) > -1; });
    });
    if (kwHits.length === 1) return { catId: kwHits[0].id, confidence: 0.92, suggest: [kwHits[0].id], reason: "Matches your keyword" };
    // 3. built-in merchant knowledge, mapped onto your category names
    var kinds = [];
    Object.keys(KINDS).forEach(function (k) {
      if (KINDS[k].m.some(function (m) { return merch.indexOf(m) > -1; })) kinds.push(k);
    });
    if (!kinds.length && tx.merchant) {
      Object.keys(KINDS).forEach(function (k) { if (KINDS[k].m.some(function (m) { return m.length > 3 && hay.indexOf(m) > -1; })) kinds.push(k); });
    }
    var cands = [];
    kinds.forEach(function (k) {
      all.forEach(function (c) { if ((c.kind || guessKind(c.name + " " + (c.group || ""))) === k && cands.indexOf(c.id) < 0) cands.push(c.id); });
    });
    kwHits.forEach(function (c) { if (cands.indexOf(c.id) < 0) cands.unshift(c.id); });
    if (cands.length === 1 && kinds.length === 1) return { catId: cands[0], confidence: 0.82, suggest: cands, reason: "Looks like " + kinds[0] };
    // Uncertain: suggest the candidates, then your most-used categories
    var freq = {};
    txs().forEach(function (t) { if (t.catId) freq[t.catId] = (freq[t.catId] || 0) + 1; });
    var popular = all.map(function (c) { return c.id; }).sort(function (a, b) { return (freq[b] || 0) - (freq[a] || 0); });
    if (near) cands = [near.catId].concat(cands.filter(function (id) { return id !== near.catId; }));
    res.suggest = cands.concat(popular.filter(function (id) { return cands.indexOf(id) < 0; })).slice(0, 4);
    res.confidence = cands.length ? 0.5 : 0.1;
    res.reason = near ? "You put “" + titleCase(near.id) + "” in " + cat(near.catId).name : cands.length ? "Could be a few categories" : "New merchant";
    return res;
  }
  function applyCategory(tx) {
    var r = categorise(tx);
    if (tx.type === "income") { tx.catId = null; tx.needsReview = false; return; }
    if (r.confidence >= 0.8) { tx.catId = r.catId; tx.needsReview = false; if (r.catId === TRANSFER) tx.type = "transfer"; }
    else { tx.catId = null; tx.needsReview = true; tx.suggest = r.suggest; }
    tx.why = r.reason;
  }

  // ======================================================================
  // Bank SMS parser — tuned for FNB, Capitec, Standard Bank, Absa, Nedbank,
  // Discovery, TymeBank, Investec, African Bank; falls back to generic rules.
  // ======================================================================
  var BANKS = [["fnb", "FNB"], ["capitec", "Capitec"], ["std bank", "Standard Bank"], ["standard bank", "Standard Bank"], ["absa", "Absa"], ["nedbank", "Nedbank"], ["discovery", "Discovery Bank"], ["tyme", "TymeBank"], ["investec", "Investec"], ["african bank", "African Bank"], ["bank zero", "Bank Zero"], ["old mutual", "Old Mutual"]];
  function parseSms(text) {
    var raw = String(text || "").trim();
    if (!raw) return { error: "Empty message" };
    var t = raw.replace(/\s+/g, " ");
    var low = t.toLowerCase();
    if (/\b(otp|one[- ]?time[- ]?pin|verification code|password|do not share)\b/.test(low)) return { error: "Looks like a one-time PIN, not a transaction" };
    if (/\b(declined|unsuccessful|failed|insufficient)\b/.test(low)) return { error: "Declined transaction — ignored" };

    // amounts: R123.45, R 1 234,56, R-45.00, ZAR123
    var re = /(^|[^A-Za-z])(?:R|ZAR)\s?(-)?\s?(\d{1,3}(?:[ ,]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)/g;
    var m, amounts = [];
    while ((m = re.exec(t))) {
      var before = low.slice(Math.max(0, m.index - 16), m.index + m[1].length);
      var isBal = /(avail|bal|balance|limit|avl|available)[^a-z]*$/.test(before);
      amounts.push({ v: parseAmount(m[3]), neg: !!m[2], bal: isBal, i: m.index });
    }
    var amt = amounts.filter(function (a) { return !a.bal && a.v > 0; })[0];
    if (!amt) return { error: "Couldn't find an amount in the message" };

    // direction
    var type = "expense";
    var incomeRe = /\b(received|deposit(ed)?|credited|inward|salary|refund|reversal|reversed|paid to you|paid into|into your|payment from|credit:|cr\b|has been paid to your|paid to (?:your )?(?:cheq|sav|a\/c|acc|account|current|s\/a))\b/;
    var expenseRe = /\b(purchase[ds]?|pur\b|spent|debit(ed)?|withdrawn|withdrawal|reserved|paid from|payment of|debit order|dr\b|cash wd|card transaction)\b/;
    if (incomeRe.test(low) && !expenseRe.test(low)) type = "income";
    if (amt.neg) type = "expense";
    var isTransfer = /\b(t\/fer|transfer(red)?|trf|tfr|eft)\b/.test(low) && !/\bpurchase|pur\b/.test(low);
    if (isTransfer && /\b(from|to) (your )?(own|savings|sav\b|cheq|current|credit card|s\/?a\b)/.test(low)) type = "transfer";

    // merchant
    var stopAhead = "(?=\\s+(?:on|from|using|with|via|for|frm|into|in|avail|ref|card|acc|a\\/c)\\b|\\s*[.;,]\\s|\\s*\\.$|\\s+\\d{1,2}[\\/\\-]\\d|\\s+\\d{1,2}\\s?(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)|\\s+\\d{1,2}:\\d{2}|$)";
    var merchant = "";
    var pats = [
      new RegExp("(?:\\bat\\s+|@\\s*)(.+?)" + stopAhead, "i"),
      /,\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s+(.+?),\s*R/i,              // Absa: "..., 12/10/26 CHECKERS SIXTY60, R-345.00"
      new RegExp("\\b(?:to|for)\\s+(.+?)" + stopAhead, "i"),
      new RegExp("\\b(?:from)\\s+(.+?)" + stopAhead, "i"),
      /(?:purchase|pur|debit order|payment)[:\s]+(?:R\s?[\d.,\s]+\s)?(.+?)(?:\.|,|$)/i
    ];
    for (var i = 0; i < pats.length && !merchant; i++) {
      var mm = t.match(pats[i]);
      if (mm && mm[1]) {
        var cand = mm[1].replace(/\bR\s?-?[\d.,\s]+/g, "").replace(/\b(cheq|cheque|savings|sav|credit card|acc|account|a\/c)[\w.\s*]*$/i, "").trim();
        cand = cand.replace(/^(of|for)\s+/i, "");
        if (cand && !/^(your|you|acc|a\/c|card|\d|\*)/i.test(cand) && !/^(the|of|a|an)$/i.test(cand) && cand.length > 1) merchant = cand;
      }
    }
    merchant = merchant.replace(/[*#]+/g, " ").replace(/\s+/g, " ").replace(/[\s.,:-]+$/, "").slice(0, 48);
    if (!merchant || (type === "income" && /^(cheq|sav|acc)/i.test(merchant))) merchant = type === "transfer" ? "Transfer" : type === "income" ? "Money received" : (/atm|cash/.test(low) ? "Cash withdrawal" : (/fee|charge/.test(low) ? "Bank fees" : "Card payment"));
    merchant = titleCase(merchant);

    // date
    var d = null, dm;
    if ((dm = t.match(/\b(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})\b/))) d = new Date(+dm[1], +dm[2] - 1, +dm[3]);
    else if ((dm = t.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/))) { var y = +dm[3]; if (y < 100) y += 2000; d = new Date(y, +dm[2] - 1, +dm[1]); }
    else if ((dm = low.match(/\b(\d{1,2})\s?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?(?:\s?(\d{4}|\d{2}(?![:\d])))?/))) {
      var yy = dm[3] ? (+dm[3] < 100 ? 2000 + +dm[3] : +dm[3]) : today().getFullYear();
      d = new Date(yy, MON3.indexOf(dm[2]), +dm[1]);
      if (!dm[3] && d > addDays(today(), 1)) d = new Date(yy - 1, d.getMonth(), d.getDate());
    }
    if (!d || isNaN(d) || d > addDays(today(), 1) || d < addDays(today(), -120)) d = null;

    var bank = "";
    BANKS.some(function (b) { if (low.indexOf(b[0]) > -1) { bank = b[1]; return true; } return false; });
    return { amount: Math.round(amt.v * 100) / 100, type: type, merchant: merchant, date: d ? iso(d) : null, bank: bank, reserved: /reserved/.test(low) };
  }
  function titleCase(s) {
    return String(s).toLowerCase().replace(/(^|[\s\-&\/(])([a-z])/g, function (_, a, b) { return a + b.toUpperCase(); })
      .replace(/\b(Fnb|Pnp|Kfc|Bp|Mtn|Atm|Dstv|Sa|Eft|Cc|Za|Mrp|Spar)\b/g, function (w) { return w.toUpperCase(); });
  }

  // Adds a transaction from SMS text. Deterministic id → no duplicates across devices.
  function ingestSms(text, at, opts) {
    var p = parseSms(text);
    if (p.error) return { ok: false, reason: p.error };
    var id = "sms_" + hash(text);
    if (db.transactions.some(function (t) { return t.id === id; })) return { ok: false, reason: "Already added", dup: true };
    var date = p.date || iso(at ? new Date(at) : new Date());
    var key = merchantKey(p.merchant);
    // FNB-style "reserved" then "paid" for the same purchase
    var twin = txs().filter(function (t) {
      return t.source === "sms" && Math.abs(t.amount - p.amount) < 0.01 && merchantKey(t.merchant) === key &&
        Math.abs(dayDiff(fromIso(t.date), fromIso(date))) <= 5 && (t.reserved || p.reserved) && t.type === p.type;
    })[0];
    if (twin) return { ok: false, reason: "Same as an earlier reserved amount", dup: true };
    var tx = touch({ id: id, date: date, amount: p.amount, type: p.type, merchant: p.merchant, note: "", source: "sms", bank: p.bank, raw: String(text).slice(0, 600), reserved: p.reserved });
    if (p.type === "transfer") { tx.catId = TRANSFER; tx.needsReview = false; } else applyCategory(tx);
    if (opts && opts.sample) tx.sample = true;
    db.transactions.push(tx);
    return { ok: true, tx: tx };
  }

  // ======================================================================
  // Rendering
  // ======================================================================
  var ICON = {
    chev: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M9 18l6-6-6-6"/></svg>',
    close: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>'
  };

  function render() {
    applyTheme();
    var p = period(ui.offset);
    $("#monthLabel").textContent = p.label;
    $("#monthSub").textContent = p.sub || (p.current ? p.left + " days left" : "");
    if (p.sub && p.current) $("#monthSub").textContent = p.sub + " · " + p.left + " days left";
    $("#monthbar").hidden = ui.tab === "budget" || ui.tab === "settings";
    var q = reviewQueue().length;
    $("#navBadge").hidden = !q; $("#navBadge").textContent = q;
    $$(".tab-btn").forEach(function (b) { b.classList.toggle("active", b.dataset.tab === ui.tab); });
    $$(".panel").forEach(function (el) { el.classList.toggle("active", el.id === "panel-" + ui.tab); });
    var fn = { home: renderHome, activity: renderActivity, insights: renderInsights, budget: renderBudget, settings: renderSettings }[ui.tab];
    $("#panel-" + ui.tab).innerHTML = fn(p);
    if (ui.tab === "insights") mountCharts(p);
    if (ui.tab === "budget") mountDrop();
    renderSyncState();
  }

  function catRow(c, spent, p) {
    var b = +c.budget || 0, fx = isFixed(c), st = paceStatus(spent, b, p, fx);
    var pct = b ? clamp(spent / b * 100, 0, 100) : (spent ? 100 : 0);
    var left = b - spent;
    var perDay = p.current && p.left > 0 && left > 0 ? money(left / p.left) + "/day" : "";
    return '<button class="cat" data-act="cat" data-id="' + c.id + '" type="button">' +
      '<span class="emoji">' + esc(c.emoji || "🏷️") + '</span><span class="body">' +
      '<span class="top"><span class="name">' + esc(c.name) + '</span><span class="amt num"><b>' + money(spent) + '</b> <span class="faint">/ ' + money(b) + '</span></span></span>' +
      '<span class="bar ' + st + '" role="progressbar" aria-valuenow="' + Math.round(pct) + '" aria-valuemin="0" aria-valuemax="100" aria-label="' + esc(c.name) + '"><span class="fill" style="width:' + pct + '%"></span>' +
      (p.current && !fx ? '<span class="pace" style="left:calc(' + (p.elapsed / p.days * 100) + '% - 1px)" title="Where you’d be if spending evenly"></span>' : "") + '</span>' +
      '<span class="meta"><span class="status ' + st + '">' + (fx && st === "ok" ? (spent >= b - 0.5 ? "✓ Paid" : spent > 0 ? "◐ Part paid" : "○ Not paid yet") : STATUS_LABEL[st]) + '</span><span class="num">' +
      (left >= 0 ? (fx && left < 0.5 ? "" : money(left) + " left" + (perDay && !fx ? " · " + perDay : "")) : money(-left) + " over") + '</span></span>' +
      '</span></button>';
  }

  function renderHome(p) {
    if (!cats().length) return onboardingHtml();
    var s = stats(p);
    var left = s.budget - s.spent;
    var pct = s.budget ? clamp(s.spent / s.budget * 100, 0, 100) : 0;
    var expected = cats().reduce(function (a, c) { return a + expectedByNow(c, s.by[c.id] || 0, p); }, 0);
    var diff = expected - s.spent;
    var st = left < 0 ? "over" : (p.current && -diff > s.budget * 0.08 ? "warn" : "ok");
    var paceAt = s.budget ? clamp(expected / s.budget * 100, 0, 100) : 0;
    var q = reviewQueue();
    var h = "";
    var sample = db.transactions.some(function (t) { return t.sample && !t.deleted; });
    if (sample) h += '<div class="card flat row small" style="background:var(--accent-soft);border-color:transparent"><span>🧪</span><span class="spacer">You’re looking at sample data.</span><button class="btn sm primary" data-act="clear-sample" type="button">Start fresh</button></div>';
    h += '<div class="grid-2"><div>';
    h += '<div class="card hero"><div class="eyebrow">' + (left >= 0 ? "Left to spend" : "Over budget") + (p.current ? "" : " · " + esc(p.label)) + '</div>' +
      '<div class="big num">' + money(Math.abs(left)) + '</div>' +
      '<div class="sub num">' + money(s.spent) + ' spent of ' + money(s.budget) + '</div>' +
      '<div class="bar ' + st + '"><span class="fill" style="width:' + pct + '%"></span>' + (p.current ? '<span class="pace" style="left:calc(' + paceAt + '% - 1px)"></span>' : "") + '</div>';
    if (p.current) {
      h += '<div class="hero-stats"><div><b class="num">' + (left > 0 && p.left > 0 ? money(left / p.left) : money(0)) + '</b><span>per day left</span></div>' +
        '<div><b class="num">' + p.left + '</b><span>days to go</span></div>' +
        '<div><b class="num">' + Math.round(pct) + '%</b><span>of budget used</span></div></div>';
      if (s.budget) h += '<div class="pace-note">' + (diff >= 0 ? "👍 <span>You’re <b>" + money(diff) + "</b> under where you’d expect to be by today.</span>" : "⚠️ <span>You’re <b>" + money(-diff) + "</b> ahead of a steady pace (bills already counted) — ease off a little.</span>") + '</div>';
    } else if (p.offset < 0) {
      h += '<div class="pace-note">' + (left >= 0 ? "🎉 You finished this month " + money(left) + " under budget." : "This month ended " + money(-left) + " over budget.") + '</div>';
    }
    h += '</div>';
    if (q.length) {
      h += '<button class="review-banner" data-act="review" type="button"><span class="emoji">🤔</span><span class="spacer"><b>' + q.length + ' transaction' + (q.length > 1 ? "s" : "") + ' to sort</b><span class="small">I wasn’t sure which category — tap to help</span></span>' + ICON.chev + '</button>';
    }
    // recent
    var recent = s.list.slice().sort(txSort).slice(0, 5);
    h += '<div class="card simple-hide"><div class="card-h"><h3>Recent</h3><button class="link" data-tab="activity" type="button">See all</button></div>';
    h += recent.length ? recent.map(txRow).join("") : '<p class="small faint" style="margin:0">No transactions yet this month. Tap + to add one, or set up SMS forwarding in Settings.</p>';
    h += '</div></div><div>';
    // categories grouped
    var gs = groups();
    gs.forEach(function (g) {
      var list = cats().filter(function (c) { return (c.group || "Expenses") === g; });
      var gSpent = list.reduce(function (a, c) { return a + (s.by[c.id] || 0); }, 0), gBud = list.reduce(function (a, c) { return a + (+c.budget || 0); }, 0);
      h += '<div class="card"><div class="card-h"><h3>' + esc(g) + '</h3><span class="small faint num">' + money(gSpent) + ' / ' + money(gBud) + '</span></div>';
      h += list.map(function (c) { return catRow(c, s.by[c.id] || 0, p); }).join("");
      h += '</div>';
    });
    if (s.uncategorised) h += '<div class="card"><button class="cat" data-act="review" type="button"><span class="emoji">❔</span><span class="body"><span class="top"><span class="name">Not sorted yet</span><span class="amt num"><b>' + money(s.uncategorised) + '</b></span></span><span class="meta"><span>Counts toward your total — tap to sort</span></span></span></button></div>';
    h += '</div></div>';
    return h;
  }

  function onboardingHtml() {
    return '<div class="card" style="padding:24px 20px"><div style="font-size:42px">👋</div><h2 style="margin:6px 0 6px;letter-spacing:-.02em">Welcome to Budgie</h2>' +
      '<p class="muted" style="margin:0 0 18px">A calm, simple way to see where your money goes each month. Let’s set up your budget — pick whatever’s easiest:</p>' +
      '<div class="stack">' +
      '<button class="btn primary block" data-act="pick-file" type="button">📄 Import my budget spreadsheet</button>' +
      '<button class="btn block" data-act="starter" type="button">✨ Start with suggested categories</button>' +
      '<button class="btn block" data-act="join" type="button">🔗 Join a family / other device</button>' +
      '<button class="btn block" data-act="sample" type="button">🧪 Just explore with sample data</button>' +
      '</div><p class="small faint" style="margin:16px 0 0">Spreadsheet tip: one column with category names and one with monthly amounts is all I need. A heading like “Fixed costs” on its own row becomes a group.</p></div>';
  }

  function txSort(a, b) { return a.date === b.date ? (b.updatedAt || 0) - (a.updatedAt || 0) : (a.date < b.date ? 1 : -1); }
  function txRow(t) {
    var c = cat(t.catId);
    var isIn = t.type === "income", isTr = t.type === "transfer" || t.catId === TRANSFER;
    var sub = isIn ? "Income" : isTr ? "Transfer · not counted" : c ? c.name : "Not sorted";
    if (t.note) sub += " · " + t.note;
    return '<button class="tx" data-act="tx" data-id="' + t.id + '" type="button"><span class="emoji">' + (isIn ? "💰" : isTr ? "🔁" : c ? esc(c.emoji) : "❔") + '</span>' +
      '<span class="body"><span class="m">' + esc(t.merchant || "Transaction") + (t.needsReview ? '<span class="pill">Sort me</span>' : "") + '</span><span class="c">' + esc(sub) + (t.source === "sms" ? " · 📩 SMS" : "") + '</span></span>' +
      '<span class="a num ' + (isIn ? "in" : isTr ? "tr" : "") + '">' + (isIn ? "+" : "−") + money(t.amount, true) + '</span></button>';
  }

  function renderActivity(p) {
    var s = stats(p);
    var qtext = ui.actQuery.toLowerCase();
    var list = s.list.filter(function (t) {
      if (ui.actFilter === "review" && !t.needsReview) return false;
      if (ui.actFilter === "income" && t.type !== "income") return false;
      if (ui.actFilter === "sms" && t.source !== "sms") return false;
      if (ui.actFilter.indexOf("cat:") === 0 && t.catId !== ui.actFilter.slice(4)) return false;
      if (qtext && (t.merchant + " " + (t.note || "") + " " + ((cat(t.catId) || {}).name || "")).toLowerCase().indexOf(qtext) < 0) return false;
      return true;
    }).sort(txSort);
    var h = '<input type="search" id="actSearch" placeholder="Search merchants, notes, categories" value="' + esc(ui.actQuery) + '">';
    var chips = [["all", "All"], ["review", "To sort (" + reviewQueue().filter(function (t) { return inPeriod(t, p); }).length + ")"], ["income", "Income"], ["sms", "From SMS"]];
    h += '<div class="chips scroll">' + chips.map(function (c) { return '<button class="chip ' + (ui.actFilter === c[0] ? "on" : "") + '" data-act="filter" data-f="' + c[0] + '" type="button">' + c[1] + '</button>'; }).join("") +
      cats().map(function (c) { var f = "cat:" + c.id; return '<button class="chip ' + (ui.actFilter === f ? "on" : "") + '" data-act="filter" data-f="' + f + '" type="button">' + esc(c.emoji) + " " + esc(c.name) + '</button>'; }).join("") + '</div>';
    if (!list.length) return h + '<div class="empty"><div class="emoji">🗒️</div>Nothing here for ' + esc(p.label) + ' yet.</div>';
    var byDay = {};
    list.forEach(function (t) { (byDay[t.date] = byDay[t.date] || []).push(t); });
    h += '<div class="card" style="padding:6px 10px 10px">';
    Object.keys(byDay).sort().reverse().forEach(function (d) {
      var dd = fromIso(d), out = byDay[d].reduce(function (a, t) { return a + (t.type === "expense" && t.catId !== TRANSFER ? t.amount : 0); }, 0);
      var label = d === iso(today()) ? "Today" : d === iso(addDays(today(), -1)) ? "Yesterday" : dd.toLocaleDateString("en-ZA", { weekday: "short", day: "numeric", month: "short" });
      h += '<div class="day-h"><span>' + label + '</span><span class="num">' + (out ? "−" + money(out, true) : "") + '</span></div>' + byDay[d].map(txRow).join("");
    });
    return h + '</div>';
  }

  // ---------- Insights ----------
  function renderInsights(p) {
    if (!cats().length) return '<div class="empty"><div class="emoji">📊</div>Set up your budget first, then insights appear here.</div>';
    var s = stats(p), prev = stats(period(p.offset - 1));
    var projected = projectSpend(p, s);
    var prevSame = p.current ? (cumulative(period(p.offset - 1))[Math.min(p.elapsed, period(p.offset - 1).days) - 1] || 0) : prev.spent;
    var noSpend = 0;
    if (p.elapsed) {
      var spentDays = {};
      s.list.forEach(function (t) { if (t.type === "expense" && t.catId !== TRANSFER) spentDays[t.date] = 1; });
      noSpend = p.elapsed - Object.keys(spentDays).filter(function (d) { return d <= iso(today()); }).length;
    }
    var income = s.income || +db.settings.income || 0;
    var saveRate = income ? (income - s.spent) / income : null;
    var vsPrev = prevSame ? (s.spent - prevSame) / prevSame : null;
    var h = '<div class="tiles four">' +
      tile(p.current ? "Heading for" : "Total spent", money(projected), projected > s.budget && s.budget ? '<span class="delta-up">▲ ' + money(projected - s.budget) + ' over budget</span>' : '<span class="delta-down">▼ ' + money(Math.max(0, s.budget - projected)) + ' under budget</span>') +
      tile(p.current ? "vs this time last month" : "vs last month", vsPrev == null ? "—" : (vsPrev > 0 ? "+" : "") + Math.round(vsPrev * 100) + "%", vsPrev == null ? "No data for last month" : money(s.spent) + " now vs " + money(prevSame)) +
      tile("Savings rate", saveRate == null ? "—" : Math.round(saveRate * 100) + "%", income ? money(income - s.spent) + " of " + money(income) + (s.income ? " income" : " expected") : "Add your income on the Budget tab") +
      tile("No-spend days", p.elapsed ? noSpend + "" : "—", p.elapsed ? "out of " + p.elapsed + " days so far" : "") + '</div>';
    h += '<div class="grid-2"><div>';
    h += '<div class="card"><div class="card-h"><h3>Spending through the month</h3></div><div class="legend"><span><i style="background:var(--series-1)"></i>' + esc(p.label) + '</span><span><i style="background:var(--series-muted)"></i>Last month</span><span><i style="background:repeating-linear-gradient(90deg,var(--ink-faint) 0 4px,transparent 4px 7px)"></i>Steady pace (bills up front)</span></div><div class="chart" id="burnChart"></div></div>';
    h += '<div class="card"><div class="card-h"><h3>Last 6 months</h3><span class="small faint">bars = spent · line = budget</span></div><div class="chart" id="monthsChart"></div></div>';
    h += '</div><div>';
    // by category
    var rows = cats().map(function (c) { return { c: c, v: s.by[c.id] || 0, pv: prev.by[c.id] || 0 }; }).sort(function (a, b) { return b.v - a.v; });
    var max = Math.max.apply(null, rows.map(function (r) { return Math.max(r.v, +r.c.budget || 0); }).concat([1]));
    h += '<div class="card"><div class="card-h"><h3>Where it went</h3><span class="small faint">marker = budget</span></div>' +
      rows.filter(function (r) { return r.v > 0 || r.c.budget; }).map(function (r) {
        return '<div class="hbar" title="' + esc(r.c.name) + ': ' + money(r.v) + ' of ' + money(r.c.budget) + '"><span class="t">' + esc(r.c.emoji) + ' ' + esc(r.c.name) + '</span><span class="v num">' + money(r.v) + ' <span class="faint" style="font-weight:500">' + (s.spent ? Math.round(r.v / s.spent * 100) + "%" : "") + '</span></span>' +
          '<span class="track"><span style="width:' + (r.v / max * 100) + '%"></span>' + (r.c.budget ? '<em style="left:' + (r.c.budget / max * 100) + '%"></em>' : "") + '</span></div>';
      }).join("") + '</div>';
    // biggest changes
    var changes = rows.map(function (r) { return { c: r.c, d: r.v - r.pv }; }).filter(function (r) { return Math.abs(r.d) >= 1; }).sort(function (a, b) { return Math.abs(b.d) - Math.abs(a.d); }).slice(0, 5);
    if (changes.length && prev.spent) {
      h += '<div class="card"><div class="card-h"><h3>Biggest changes vs last month</h3></div>' + changes.map(function (r) {
        return '<div class="summary-line"><span>' + esc(r.c.emoji) + ' ' + esc(r.c.name) + '</span><span class="num ' + (r.d > 0 ? "delta-up" : "delta-down") + '">' + (r.d > 0 ? "▲ " : "▼ ") + money(Math.abs(r.d)) + '</span></div>';
      }).join("") + '</div>';
    }
    // top merchants
    var mm = {};
    s.list.forEach(function (t) { if (t.type !== "expense" || t.catId === TRANSFER) return; var k = t.merchant || "Other"; mm[k] = mm[k] || { n: 0, v: 0 }; mm[k].n++; mm[k].v += t.amount; });
    var top = Object.keys(mm).map(function (k) { return { k: k, n: mm[k].n, v: mm[k].v }; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 6);
    if (top.length) h += '<div class="card"><div class="card-h"><h3>Top places</h3></div>' + top.map(function (r) { return '<div class="summary-line"><span>' + esc(r.k) + ' <span class="faint small">× ' + r.n + '</span></span><span class="num"><b>' + money(r.v) + '</b></span></div>'; }).join("") + '</div>';
    // biggest single
    var big = s.list.filter(function (t) { return t.type === "expense" && t.catId !== TRANSFER; }).sort(function (a, b) { return b.amount - a.amount; }).slice(0, 3);
    if (big.length) h += '<div class="card"><div class="card-h"><h3>Biggest purchases</h3></div>' + big.map(txRow).join("") + '</div>';
    h += '</div></div>';
    return h;
  }
  // Fixed bills count at their budget (or actual if more); variable spending is extrapolated.
  function projectSpend(p, s) {
    if (!p.current || !p.elapsed) return s.spent;
    var fixed = 0, variable = s.uncategorised;
    cats().forEach(function (c) { var v = s.by[c.id] || 0; if (isFixed(c)) fixed += Math.max(v, +c.budget || 0); else variable += v; });
    return fixed + variable / p.elapsed * p.days;
  }
  function tile(label, v, d) { return '<div class="tile"><div class="eyebrow">' + label + '</div><div class="v num">' + v + '</div><div class="d">' + d + '</div></div>'; }

  function mountCharts(p) {
    burnChart($("#burnChart"), p);
    monthsChart($("#monthsChart"), p);
  }
  function cumulative(p) {
    var daily = new Array(p.days).fill(0);
    txs().forEach(function (t) {
      if (!inPeriod(t, p) || t.type !== "expense" || t.catId === TRANSFER) return;
      var i = dayDiff(p.start, fromIso(t.date)); if (i >= 0 && i < p.days) daily[i] += t.amount;
    });
    var run = 0; return daily.map(function (v) { run += v; return run; });
  }
  // max such that max/4 is a round step (ticks like R10k, R20k...)
  function niceMax(v) { if (v <= 0) return 100; var st = v / 4, e = Math.pow(10, Math.floor(Math.log10(st))), f = st / e; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * e * 4; }
  function shortMoney(v) { return v >= 1000 ? "R" + (v / 1000).toFixed(v >= 10000 ? 0 : 1).replace(/\.0$/, "") + "k" : "R" + Math.round(v); }

  function burnChart(el, p) {
    if (!el) return;
    var budget = stats(p).budget, cur = cumulative(p), prevP = period(p.offset - 1), prev = cumulative(prevP);
    var shown = p.current ? cur.slice(0, p.elapsed) : cur;
    var W = Math.max(300, el.clientWidth || 560), H = Math.round(clamp(W * 0.5, 190, 260)), L = 46, R = 10, T = 12, B = 26;
    var max = niceMax(Math.max(budget, cur[cur.length - 1] || 0, prev[prev.length - 1] || 0));
    var fixedBudget = cats().reduce(function (a, c) { return a + (isFixed(c) ? +c.budget || 0 : 0); }, 0);
    var paceAt = function (i) { return fixedBudget + (budget - fixedBudget) * (n === 1 ? 1 : i / (n - 1)); };
    var n = p.days;
    var x = function (i) { return L + (W - L - R) * (n === 1 ? 0 : i / (n - 1)); };
    var y = function (v) { return T + (H - T - B) * (1 - v / max); };
    var path = function (arr) { return arr.map(function (v, i) { return (i ? "L" : "M") + x(i).toFixed(1) + "," + y(v).toFixed(1); }).join(""); };
    var g = "";
    for (var k = 0; k <= 4; k++) { var v = max * k / 4; g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="var(--grid)" stroke-width="1"/><text class="axis num" x="' + (L - 6) + '" y="' + (y(v) + 3) + '" text-anchor="end">' + shortMoney(v) + '</text>'; }
    [0, Math.floor((n - 1) / 2), n - 1].forEach(function (i) { var d = addDays(p.start, i); g += '<text class="axis" x="' + x(i) + '" y="' + (H - 8) + '" text-anchor="' + (i === 0 ? "start" : i === n - 1 ? "end" : "middle") + '">' + d.getDate() + " " + MONTHS[d.getMonth()].slice(0, 3) + '</text>'; });
    var prevScaled = prev.slice(0, n);
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Cumulative spending this month vs last month and an even pace">' + g +
      (budget ? '<line x1="' + x(0) + '" y1="' + y(paceAt(0)) + '" x2="' + x(n - 1) + '" y2="' + y(budget) + '" stroke="var(--ink-faint)" stroke-width="1.5" stroke-dasharray="4 4"/>' : "") +
      '<path d="' + path(prevScaled) + '" fill="none" stroke="var(--series-muted)" stroke-width="2" stroke-linejoin="round"/>' +
      '<path d="' + path(shown) + '" fill="none" stroke="var(--series-1)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>' +
      (shown.length ? '<circle cx="' + x(shown.length - 1) + '" cy="' + y(shown[shown.length - 1]) + '" r="4.5" fill="var(--series-1)" stroke="var(--surface)" stroke-width="2"/>' : "") +
      '<line id="bcX" y1="' + T + '" y2="' + (H - B) + '" stroke="var(--ink-faint)" stroke-width="1" visibility="hidden"/>' +
      '<rect x="' + L + '" y="' + T + '" width="' + (W - L - R) + '" height="' + (H - T - B) + '" fill="transparent" id="bcHit"/></svg><div class="tip" hidden></div>';
    el.innerHTML = svg;
    var hit = $("#bcHit", el), tip = $(".tip", el), xl = $("#bcX", el), svgEl = $("svg", el);
    function move(ev) {
      var r = svgEl.getBoundingClientRect(); var px = ((ev.touches ? ev.touches[0].clientX : ev.clientX) - r.left) / r.width * W;
      var i = clamp(Math.round((px - L) / (W - L - R) * (n - 1)), 0, n - 1);
      xl.setAttribute("x1", x(i)); xl.setAttribute("x2", x(i)); xl.setAttribute("visibility", "visible");
      var d = addDays(p.start, i);
      tip.hidden = false;
      tip.style.left = (x(i) / W * 100) + "%"; tip.style.top = (y(Math.max(shown[i] || 0, prevScaled[i] || 0)) / H * 100) + "%";
      tip.innerHTML = "<b>" + d.getDate() + " " + MONTHS[d.getMonth()].slice(0, 3) + "</b>" +
        (i < shown.length ? "This month: " + money(shown[i]) + "<br>" : "") + "Last month: " + money(prevScaled[i] || 0) + (budget ? "<br>Steady pace: " + money(paceAt(i)) : "");
    }
    function leave() { tip.hidden = true; xl.setAttribute("visibility", "hidden"); }
    hit.addEventListener("mousemove", move); hit.addEventListener("touchmove", move, { passive: true }); hit.addEventListener("touchstart", move, { passive: true });
    hit.addEventListener("mouseleave", leave); hit.addEventListener("touchend", leave);
  }

  function monthsChart(el, p) {
    if (!el) return;
    var ms = [];
    for (var i = 5; i >= 0; i--) { var pp = period(p.offset - i); var st = stats(pp); ms.push({ p: pp, v: st.spent, b: st.budget }); }
    var W = Math.max(300, el.clientWidth || 560), H = Math.round(clamp(W * 0.45, 170, 230)), L = 46, R = 10, T = 12, B = 26;
    var max = niceMax(Math.max.apply(null, ms.map(function (m) { return Math.max(m.v, m.b); }).concat([1])));
    var bw = (W - L - R) / ms.length;
    var y = function (v) { return T + (H - T - B) * (1 - v / max); };
    var g = "";
    for (var k = 0; k <= 4; k++) { var v = max * k / 4; g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="var(--grid)"/><text class="axis num" x="' + (L - 6) + '" y="' + (y(v) + 3) + '" text-anchor="end">' + shortMoney(v) + '</text>'; }
    var bars = ms.map(function (m, i) {
      var x0 = L + i * bw + bw * 0.22, w = bw * 0.56, top = y(m.v), hh = Math.max(0, y(0) - top);
      var over = m.b && m.v > m.b;
      return '<g class="mbar" data-i="' + i + '"><rect x="' + (L + i * bw) + '" y="' + T + '" width="' + bw + '" height="' + (H - T - B) + '" fill="transparent"/>' +
        (hh > 0 ? '<path d="M' + x0 + ',' + y(0) + 'V' + (top + Math.min(4, hh)) + 'q0,-4 4,-4h' + (w - 8) + 'q4,0 4,4V' + y(0) + 'z" fill="' + (over ? "var(--over)" : "var(--series-1)") + '" opacity="' + (m.p.offset === p.offset ? 1 : 0.75) + '"/>' : "") +
        (m.b ? '<line x1="' + (x0 - 4) + '" x2="' + (x0 + w + 4) + '" y1="' + y(m.b) + '" y2="' + y(m.b) + '" stroke="var(--ink)" stroke-width="2" stroke-linecap="round"/>' : "") +
        '<text class="axis" x="' + (x0 + w / 2) + '" y="' + (H - 8) + '" text-anchor="middle" font-weight="' + (m.p.offset === p.offset ? 700 : 400) + '">' + m.p.label.slice(0, 3) + '</text></g>';
    }).join("");
    el.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Total spent per month for the last six months, with budget markers">' + g + bars + '</svg><div class="tip" hidden></div>';
    var tip = $(".tip", el);
    $$(".mbar", el).forEach(function (gEl) {
      var m = ms[+gEl.dataset.i];
      var show = function () {
        tip.hidden = false; tip.style.left = ((L + (+gEl.dataset.i + 0.5) * bw) / W * 100) + "%"; tip.style.top = (y(Math.max(m.v, m.b)) / H * 100) + "%";
        tip.innerHTML = "<b>" + esc(m.p.label) + "</b>Spent " + money(m.v) + (m.b ? "<br>Budget " + money(m.b) : "") + (m.b ? "<br>" + (m.v > m.b ? "▲ " + money(m.v - m.b) + " over" : "▼ " + money(m.b - m.v) + " under") : "");
      };
      gEl.addEventListener("mouseenter", show); gEl.addEventListener("touchstart", show, { passive: true });
      gEl.addEventListener("mouseleave", function () { tip.hidden = true; });
    });
  }

  // ---------- Budget tab ----------
  function renderBudget() {
    var all = cats();
    var total = all.reduce(function (a, c) { return a + (+c.budget || 0); }, 0);
    var inc = +db.settings.income || 0;
    var h = '<div class="grid-2"><div>';
    h += '<div class="card"><div class="card-h"><h3>Monthly plan</h3></div>' +
      '<label class="field">Expected monthly income (take-home)<input id="incomeInput" inputmode="decimal" value="' + (inc ? inc : "") + '" placeholder="e.g. 32 000"></label>' +
      '<div style="margin-top:12px"><div class="summary-line"><span class="muted">Income</span><span class="num">' + money(inc) + '</span></div>' +
      '<div class="summary-line"><span class="muted">Budgeted across ' + all.length + ' categories</span><span class="num">− ' + money(total) + '</span></div>' +
      '<div class="summary-line total"><span>' + (inc - total >= 0 ? "Not yet given a job" : "Over-planned by") + '</span><span class="num ' + (inc - total >= 0 ? "delta-down" : "delta-up") + '">' + money(Math.abs(inc - total)) + '</span></div></div>' +
      (inc && inc - total > 0 ? '<p class="small faint" style="margin:8px 0 0">Tip: give every rand a job — put the extra into savings or a goal category.</p>' : "") + '</div>';
    h += '<div class="drop" id="dropZone"><div class="emoji">📄</div><b>Drop an Excel or CSV budget here</b><p class="small muted" style="margin:4px 0 12px">Columns like <i>Category</i> and <i>Amount</i>. You’ll get a preview before anything changes.</p>' +
      '<div class="row" style="justify-content:center;gap:8px" ><button class="btn sm primary" data-act="pick-file" type="button">Choose file</button><button class="btn sm" data-act="template" type="button">Get a template</button></div></div>';
    h += '</div><div>';
    h += '<div class="card"><div class="card-h"><h3>Categories</h3><button class="btn sm primary" data-act="new-cat" type="button">+ Add</button></div>';
    if (!all.length) h += '<p class="small faint">No categories yet. Import a spreadsheet, or <button class="link" data-act="starter" type="button">add suggested ones</button>.</p>';
    groups().forEach(function (g) {
      var list = all.filter(function (c) { return (c.group || "Expenses") === g; });
      h += '<div class="group-h"><span class="eyebrow">' + esc(g) + '</span><span class="small faint num">' + money(list.reduce(function (a, c) { return a + (+c.budget || 0); }, 0)) + '</span></div>';
      h += list.map(function (c) { return '<button class="plan-row" data-act="cat-edit" data-id="' + c.id + '" type="button"><span class="emoji">' + esc(c.emoji) + '</span><span class="n">' + esc(c.name) + '</span><span class="v num">' + money(c.budget) + '</span>' + ICON.chev + '</button>'; }).join("");
    });
    h += '</div></div></div>';
    return h;
  }

  // ---------- Settings ----------
  function renderSettings() {
    var url = local.syncUrl || "", code = local.syncCode || "";
    var h = '<div class="grid-2"><div>';
    h += '<div class="card"><div class="card-h"><h3>Sync & bank SMS</h3><span class="small ' + (code && url ? "delta-down" : "faint") + '">' + (code && url ? "● On" : "Off") + '</span></div>';
    if (code && url) {
      h += '<p class="small muted" style="margin-top:0">This device syncs with everyone using the same code. Data is encrypted on your device before it’s uploaded.</p>' +
        '<div class="small muted">Your sync code</div><div class="copy-box"><span>' + esc(code) + '</span><button class="btn sm" data-copy="' + esc(code) + '" type="button">Copy</button></div>' +
        '<div class="small muted" style="margin-top:10px">Invite link for another device (opens Budgie and joins)</div><div class="copy-box"><span>' + esc(joinLink()) + '</span><button class="btn sm" data-copy="' + esc(joinLink()) + '" type="button">Copy</button></div>' +
        '<p class="small faint">Last synced: ' + (local.lastSync ? new Date(local.lastSync).toLocaleString("en-ZA") : "never") + '</p>' +
        '<div class="row wrap"><button class="btn sm primary" data-act="sync-now" type="button">Sync now</button><button class="btn sm" data-act="sms-guide" type="button">📩 Set up bank SMS</button><button class="btn sm danger" data-act="sync-off" type="button">Turn off on this device</button></div>';
    } else {
      h += '<p class="small muted" style="margin-top:0">Keep your phone, laptop and family devices in sync, and let bank SMSes flow in automatically.</p>' +
        '<label class="field">Sync server URL<input id="syncUrl" placeholder="https://budgie-sync.yourname.workers.dev" value="' + esc(url) + '"></label>' +
        '<div class="row wrap" style="margin-top:10px"><button class="btn sm primary" data-act="sync-create" type="button">Create new sync space</button><button class="btn sm" data-act="join" type="button">Join with a code</button></div>';
    }
    h += '</div>';
    h += '<div class="card"><div class="card-h"><h3>Display</h3></div><div class="stack">' +
      '<label class="check"><input type="checkbox" id="simpleToggle" ' + (local.simple ? "checked" : "") + '> <span><b>Easy view</b><br><span class="small muted">Bigger text and fewer screens — great for parents. This device only.</span></span></label>' +
      '<div><div class="small muted" style="margin-bottom:6px">Theme</div><div class="seg" id="themeSeg">' + ["auto", "light", "dark"].map(function (t) { return '<button type="button" data-theme-v="' + t + '" class="' + (local.theme === t ? "on" : "") + '">' + t[0].toUpperCase() + t.slice(1) + '</button>'; }).join("") + '</div></div>' +
      '</div></div>';
    h += '<div class="card"><div class="card-h"><h3>Budget month</h3></div>' +
      '<label class="field">My budget month starts on day (payday)<select id="startDay">' + Array.from({ length: 28 }, function (_, i) { return '<option ' + (+db.settings.startDay === i + 1 ? "selected" : "") + '>' + (i + 1) + '</option>'; }).join("") + '</select></label>' +
      '<p class="small faint" style="margin:8px 0 0">Paid on the 25th? Set 25 and each budget month runs 25th → 24th.</p>' +
      '<label class="field" style="margin-top:12px">Currency<select id="currency">' + ["ZAR", "USD", "GBP", "EUR", "NAD", "BWP"].map(function (c) { return '<option ' + (db.settings.currency === c ? "selected" : "") + '>' + c + '</option>'; }).join("") + '</select></label></div>';
    h += '<div class="card"><div class="card-h"><h3>Test the SMS reader</h3></div><p class="small muted" style="margin-top:0">Paste one or more bank SMSes (one per line). I’ll add them as transactions.</p>' +
      '<textarea id="smsPaste" placeholder="FNB :-) R245.50 paid from cheq a/c..123456 @ Checkers Hyper. Avail R4 321.00. 14Oct 13:02"></textarea>' +
      '<div class="row" style="margin-top:10px"><button class="btn sm" data-act="sms-preview" type="button">Preview</button><button class="btn sm primary" data-act="sms-add" type="button">Add transactions</button></div><div id="smsResult" class="stack small" style="margin-top:10px"></div></div>';
    h += '</div><div>';
    h += '<div class="card"><div class="card-h"><h3>Your data</h3></div><div class="row wrap">' +
      '<button class="btn sm" data-act="export" type="button">⬇️ Download backup</button><button class="btn sm" data-act="import-backup" type="button">⬆️ Restore backup</button>' +
      '<button class="btn sm" data-act="export-csv" type="button">Transactions as CSV</button></div>' +
      '<p class="small faint">Learned merchant rules: ' + db.rules.filter(function (r) { return !r.deleted; }).length + ' · Transactions: ' + txs().length + '</p>' +
      '<button class="btn sm danger" data-act="wipe" type="button">Erase everything on this device</button></div>';
    h += '<p class="small faint" style="text-align:center">Budgie · made for you · v1.0</p>';
    h += '</div></div>';
    return h;
  }
  function joinLink() { return location.origin + location.pathname + "#join=" + encodeURIComponent(local.syncCode) + "&server=" + encodeURIComponent(local.syncUrl); }

  // ======================================================================
  // Sheets
  // ======================================================================
  function openSheet(html, onMount) {
    $("#sheet").innerHTML = '<div class="grab"></div>' + html;
    $("#sheetWrap").hidden = false;
    $("#sheet").scrollTop = 0;
    if (onMount) onMount($("#sheet"));
  }
  function closeSheet() { $("#sheetWrap").hidden = true; $("#sheet").innerHTML = ""; ui.sheet = null; }
  function sheetHead(title, sub) { return '<div class="sheet-h"><div><h2>' + title + '</h2>' + (sub ? '<div class="small muted">' + sub + '</div>' : "") + '</div><button class="icon-btn" data-close type="button" aria-label="Close">' + ICON.close + '</button></div>'; }

  function catPicker(selected, suggest) {
    suggest = suggest || [];
    return '<div class="cat-grid">' + cats().map(function (c) {
      return '<button type="button" class="cat-pick ' + (selected === c.id ? "on" : suggest.indexOf(c.id) > -1 && !selected ? "suggest" : "") + '" data-pick="' + c.id + '"><span class="e">' + esc(c.emoji) + '</span>' + esc(c.name) + '</button>';
    }).join("") + '</div>';
  }

  function txSheet(t) {
    var isNew = !t;
    t = t || { id: uid(), date: iso(today()), amount: "", type: "expense", merchant: "", catId: null, note: "", source: "manual" };
    if (isNew && ui.offset !== 0) { var pp = period(ui.offset); t.date = pp.e < iso(today()) ? pp.e : pp.s; }
    var state = { type: t.type === "transfer" ? "transfer" : t.type, catId: t.catId === TRANSFER ? null : t.catId };
    var html = sheetHead(isNew ? "Add transaction" : "Edit transaction", t.source === "sms" ? "From " + (t.bank || "bank") + " SMS" : "") +
      '<div class="seg" id="typeSeg">' + [["expense", "Spent"], ["income", "Received"], ["transfer", "Transfer"]].map(function (x) { return '<button type="button" data-type="' + x[0] + '" class="' + (state.type === x[0] ? "on" : "") + '">' + x[1] + '</button>'; }).join("") + '</div>' +
      '<input class="amount-input num" id="txAmount" inputmode="decimal" placeholder="R 0" value="' + (t.amount ? String(t.amount).replace(".", ",") : "") + '" autocomplete="off">' +
      '<div class="stack"><label class="field">Where / what<input id="txMerchant" value="' + esc(t.merchant) + '" placeholder="e.g. Woolworths" autocomplete="off"></label>' +
      '<div id="catWrap"><div class="small muted" style="margin-bottom:6px;font-weight:600">Category</div>' + catPicker(state.catId, t.suggest) + '</div>' +
      '<div class="row"><label class="field" style="flex:1">Date<input type="date" id="txDate" value="' + t.date + '"></label><label class="field" style="flex:1.4">Note<input id="txNote" value="' + esc(t.note || "") + '" placeholder="Optional"></label></div>' +
      (t.raw ? '<div class="raw">' + esc(t.raw) + '</div>' : "") +
      '<label class="check" id="rememberWrap"><input type="checkbox" id="txRemember" ' + (t.source === "sms" ? "checked" : "") + '> Always use this category for this merchant</label>' +
      '<button class="btn primary block sticky-save" id="txSave" type="button">' + (isNew ? "Add" : "Save") + '</button>' +
      (isNew ? "" : '<button class="btn danger block" id="txDelete" type="button">Delete</button>') + '</div>';
    openSheet(html, function (el) {
      var catWrap = $("#catWrap", el);
      function syncType() { catWrap.hidden = state.type !== "expense"; $("#rememberWrap", el).hidden = state.type !== "expense"; $$("#typeSeg button", el).forEach(function (b) { b.classList.toggle("on", b.dataset.type === state.type); }); }
      syncType();
      $$("#typeSeg button", el).forEach(function (b) { b.onclick = function () { state.type = b.dataset.type; syncType(); }; });
      catWrap.addEventListener("click", function (e) {
        var b = e.target.closest("[data-pick]"); if (!b) return;
        state.catId = b.dataset.pick; $$("[data-pick]", catWrap).forEach(function (x) { x.classList.toggle("on", x === b); x.classList.remove("suggest"); });
      });
      if (isNew) setTimeout(function () { $("#txAmount", el).focus(); }, 250);
      $("#txSave", el).onclick = function () {
        var amt = Math.abs(parseAmount($("#txAmount", el).value));
        if (!amt) { toast("Enter an amount"); $("#txAmount", el).focus(); return; }
        var existing = db.transactions.filter(function (x) { return x.id === t.id; })[0];
        var rec = existing || t;
        rec.amount = Math.round(amt * 100) / 100; rec.type = state.type;
        rec.merchant = $("#txMerchant", el).value.trim() || (state.type === "income" ? "Income" : "Transaction");
        rec.date = $("#txDate", el).value || iso(today()); rec.note = $("#txNote", el).value.trim();
        if (state.type === "expense") {
          rec.catId = state.catId;
          if (!rec.catId && isNew) { applyCategory(rec); } else rec.needsReview = !rec.catId;
          if (rec.catId && $("#txRemember", el).checked) learn(rec.merchant, rec.catId);
        } else { rec.catId = state.type === "transfer" ? TRANSFER : null; rec.needsReview = false; }
        delete rec.suggest; touch(rec);
        if (!existing) db.transactions.push(rec);
        closeSheet(); commit();
        toast(isNew ? "Added " + money(rec.amount, true) + (rec.catId && cat(rec.catId) ? " to " + cat(rec.catId).name : "") : "Saved");
      };
      if (!isNew) $("#txDelete", el).onclick = function () {
        var rec = db.transactions.filter(function (x) { return x.id === t.id; })[0];
        rec.deleted = true; touch(rec); closeSheet(); commit();
        toast("Deleted", "Undo", function () { rec.deleted = false; touch(rec); commit(); });
      };
    });
  }

  function learn(merchant, catId) {
    var key = merchantKey(merchant); if (!key) return 0;
    var r = db.rules.filter(function (x) { return x.id === key; })[0];
    if (r) { r.catId = catId; r.deleted = false; touch(r); } else db.rules.push(touch({ id: key, catId: catId, label: merchant }));
    // sort other waiting transactions from the same merchant
    var n = 0;
    txs().forEach(function (t) { if (t.needsReview && merchantKey(t.merchant) === key) { t.catId = catId; t.needsReview = false; if (catId === TRANSFER) t.type = "transfer"; delete t.suggest; touch(t); n++; } });
    return n;
  }

  // Review flow: one card at a time, suggestions first (like Copilot's review queue)
  function reviewSheet() {
    var q = reviewQueue();
    if (!q.length) { closeSheet(); toast("All sorted 🎉"); return; }
    var t = q[0];
    var sug = (t.suggest && t.suggest.length ? t.suggest : categorise(t).suggest).filter(function (id) { return cat(id); });
    var html = sheetHead("Which category?", q.length > 1 ? "1 of " + q.length + " to sort" : "Last one!") +
      '<div style="text-align:center;margin:4px 0 14px"><div class="small muted">' + fromIso(t.date).toLocaleDateString("en-ZA", { weekday: "long", day: "numeric", month: "long" }) + '</div>' +
      '<div style="font-size:2.1em;font-weight:800;letter-spacing:-.02em" class="num">' + money(t.amount, true) + '</div><div style="font-weight:700;font-size:1.05em">' + esc(t.merchant) + '</div>' +
      (t.why ? '<div class="small faint">' + esc(t.why) + '</div>' : "") + '</div>' +
      (t.raw ? '<div class="raw" style="margin-bottom:12px">' + esc(t.raw) + '</div>' : "") +
      (sug.length ? '<div class="small muted" style="font-weight:600;margin-bottom:6px">Best guesses</div><div class="chips" style="margin-bottom:12px">' + sug.map(function (id) { var c = cat(id); return '<button type="button" class="chip suggest" data-pick="' + id + '">' + esc(c.emoji) + " " + esc(c.name) + '</button>'; }).join("") + '</div>' : "") +
      '<label class="check" style="margin-bottom:10px"><input type="checkbox" id="rvRemember" checked> Remember for “' + esc(titleCase(merchantKey(t.merchant)) || t.merchant) + '” next time</label>' +
      '<div class="row wrap" style="margin-bottom:16px"><button class="btn sm" data-rv="transfer" type="button">🔁 Not spending</button><button class="btn sm" data-rv="skip" type="button">Later</button><span class="spacer"></span><button class="btn sm" data-rv="edit" type="button">Edit…</button></div>' +
      '<div class="small muted" style="font-weight:600;margin-bottom:6px">All categories</div>' + catPicker(null, []);
    openSheet(html, function (el) {
      el.onclick = function (e) {
        var pick = e.target.closest("[data-pick]"), rv = e.target.closest("[data-rv]");
        if (pick) resolve(pick.dataset.pick);
        else if (rv && rv.dataset.rv === "transfer") resolve(TRANSFER);
        else if (rv && rv.dataset.rv === "skip") { closeSheet(); }
        else if (rv && rv.dataset.rv === "edit") { el.onclick = null; txSheet(t); }
      };
      function resolve(catId) {
        el.onclick = null;
        var remember = $("#rvRemember", el).checked;
        t.catId = catId; t.needsReview = false; delete t.suggest; if (catId === TRANSFER) t.type = "transfer"; touch(t);
        var more = remember ? learn(t.merchant, catId) : 0;
        persist(); scheduleSync(); render();
        if (more) toast("Also sorted " + more + " more from " + t.merchant);
        reviewSheet();
      }
    });
  }

  function catSheet(c) {
    var isNew = !c;
    c = c || { id: uid(), name: "", emoji: "🏷️", budget: "", group: groups()[0] || "Expenses", keywords: "" };
    var emojis = ["🏠", "🛒", "🍔", "⛽", "🚕", "🚗", "💡", "📱", "🛡️", "💊", "📺", "🛍️", "🎉", "💇", "🏋️", "🎓", "🏦", "🎁", "🐾", "🏧", "💵", "💳", "👶", "✈️", "🧹", "☕", "🍷", "📚", "🧾", "🏷️"];
    var html = sheetHead(isNew ? "New category" : "Edit category") + '<div class="stack">' +
      '<div class="row"><button type="button" class="emoji" id="emojiBtn" style="width:52px;height:52px;font-size:26px;border:none">' + esc(c.emoji) + '</button><label class="field" style="flex:1">Name<input id="cName" value="' + esc(c.name) + '" placeholder="e.g. Groceries"></label></div>' +
      '<div class="chips" id="emojiGrid" hidden>' + emojis.map(function (e) { return '<button class="chip" type="button" data-emoji="' + e + '" style="font-size:20px;padding:6px 10px">' + e + '</button>'; }).join("") + '</div>' +
      '<label class="field">Monthly budget<input id="cBudget" inputmode="decimal" value="' + (c.budget || "") + '" placeholder="R 0"></label>' +
      '<label class="field">Group<input id="cGroup" list="groupList" value="' + esc(c.group || "") + '" placeholder="e.g. Fixed costs"><datalist id="groupList">' + groups().map(function (g) { return '<option value="' + esc(g) + '">'; }).join("") + '</datalist></label>' +
      '<label class="check"><input type="checkbox" id="cFixed" ' + (isFixed(c) ? "checked" : "") + '> <span><b>Fixed bill</b><br><span class="small muted">Paid once a month (rent, insurance, debit orders) — not paced day by day</span></span></label>' +
      '<label class="field">Auto-sort keywords <span class="faint" style="font-weight:400">— comma separated, matched against SMS text</span><input id="cKw" value="' + esc(c.keywords || "") + '" placeholder="e.g. checkers, spar"></label>' +
      '<button class="btn primary block" id="cSave" type="button">' + (isNew ? "Add category" : "Save") + '</button>' +
      (isNew ? "" : '<button class="btn danger block" id="cDel" type="button">Delete category</button>') + '</div>';
    openSheet(html, function (el) {
      var nameIn = $("#cName", el);
      nameIn.addEventListener("input", function () { if (isNew && (c.emoji === "🏷️" || c._auto)) { c.emoji = emojiFor(nameIn.value); c._auto = true; $("#emojiBtn", el).textContent = c.emoji; } });
      $("#emojiBtn", el).onclick = function () { $("#emojiGrid", el).hidden = !$("#emojiGrid", el).hidden; };
      $("#emojiGrid", el).onclick = function (e) { var b = e.target.closest("[data-emoji]"); if (!b) return; c.emoji = b.dataset.emoji; c._auto = false; $("#emojiBtn", el).textContent = c.emoji; $("#emojiGrid", el).hidden = true; };
      if (isNew) setTimeout(function () { nameIn.focus(); }, 250);
      $("#cSave", el).onclick = function () {
        var name = nameIn.value.trim(); if (!name) { toast("Give it a name"); return; }
        var rec = cat(c.id) || { id: c.id, order: cats().length };
        rec.name = name; rec.emoji = c.emoji; rec.budget = Math.abs(parseAmount($("#cBudget", el).value)) || 0;
        rec.group = $("#cGroup", el).value.trim() || "Expenses"; rec.keywords = $("#cKw", el).value.trim(); rec.kind = guessKind(name + " " + rec.group);
        var autoFixed = isFixed({ name: rec.name, group: rec.group, kind: rec.kind });
        rec.fixed = $("#cFixed", el).checked === autoFixed ? null : $("#cFixed", el).checked;
        touch(rec); if (!cat(c.id)) db.categories.push(rec);
        closeSheet(); commit(); toast(isNew ? "Added " + name : "Saved");
      };
      if (!isNew) $("#cDel", el).onclick = function () {
        var used = txs().filter(function (t) { return t.catId === c.id; }).length;
        var btn = $("#cDel", el);
        if (!btn.dataset.armed) { btn.dataset.armed = 1; btn.textContent = used ? "Tap again — " + used + " transactions become unsorted" : "Tap again to delete"; return; }
        var rec = cat(c.id); rec.deleted = true; touch(rec);
        txs().forEach(function (t) { if (t.catId === c.id) { t.catId = null; t.needsReview = true; touch(t); } });
        closeSheet(); commit(); toast("Deleted " + rec.name);
      };
    });
  }

  function catDetailSheet(c) {
    var p = period(ui.offset), s = stats(p), spent = s.by[c.id] || 0;
    var list = s.list.filter(function (t) { return t.catId === c.id; }).sort(txSort);
    var st = paceStatus(spent, c.budget, p, isFixed(c));
    var html = sheetHead(esc(c.emoji) + " " + esc(c.name), esc(p.label)) +
      '<div class="tiles" style="margin-bottom:12px">' + tile("Spent", money(spent), '<span class="status ' + st + '">' + STATUS_LABEL[st] + '</span>') + tile(c.budget - spent >= 0 ? "Left" : "Over", money(Math.abs(c.budget - spent)), "of " + money(c.budget) + " budget") + '</div>' +
      '<div class="bar ' + st + '" style="margin-bottom:14px"><span class="fill" style="width:' + (c.budget ? clamp(spent / c.budget * 100, 0, 100) : 0) + '%"></span>' + (p.current && !isFixed(c) ? '<span class="pace" style="left:calc(' + (p.elapsed / p.days * 100) + '% - 1px)"></span>' : "") + '</div>' +
      (list.length ? list.map(txRow).join("") : '<p class="small faint">Nothing spent here yet.</p>') +
      '<div class="row" style="margin-top:14px"><button class="btn block" id="cdEdit" type="button">Edit category & budget</button></div>';
    openSheet(html, function (el) {
      $("#cdEdit", el).onclick = function () { catSheet(c); };
      el.onclick = function (e) { var b = e.target.closest('[data-act="tx"]'); if (b) { el.onclick = null; txSheet(db.transactions.filter(function (t) { return t.id === b.dataset.id; })[0]); } };
    });
  }

  // ======================================================================
  // Spreadsheet import
  // ======================================================================
  function loadSheetJS() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise(function (res, rej) {
      var s = document.createElement("script"); s.src = SHEETJS_URL;
      s.onload = function () { res(window.XLSX); }; s.onerror = function () { rej(new Error("Couldn't load the Excel reader — are you online?")); };
      document.head.appendChild(s);
    });
  }
  function parseCsv(text) {
    var rows = [], row = [], cell = "", q = false;
    var delim = (text.split("\n")[0].match(/;/g) || []).length > (text.split("\n")[0].match(/,/g) || []).length ? ";" : ",";
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ""; }
      else if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cell); rows.push(row); row = []; cell = ""; }
      else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  function handleFile(file) {
    if (!file) return;
    var name = file.name.toLowerCase();
    if (/\.json$/.test(name)) { restoreBackup(file); return; }
    var reader = new FileReader();
    if (/\.csv$|\.txt$/.test(name)) {
      reader.onload = function () { importRows(parseCsv(String(reader.result)), file.name); };
      reader.readAsText(file);
    } else {
      toast("Reading " + file.name + "…");
      loadSheetJS().then(function (X) {
        reader.onload = function () {
          try {
            var wb = X.read(new Uint8Array(reader.result), { type: "array" });
            // choose the sheet that looks most like a budget
            var best = null, bestScore = -1;
            wb.SheetNames.forEach(function (sn) {
              var rows = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: "" });
              var score = detectBudget(rows).items.length + (/budget/i.test(sn) ? 5 : 0);
              if (score > bestScore) { best = rows; bestScore = score; }
            });
            importRows(best || [], file.name);
          } catch (e) { toast("Couldn't read that file: " + e.message); }
        };
        reader.readAsArrayBuffer(file);
      }).catch(function (e) { toast(e.message); });
    }
  }

  // Figures out which columns are names / amounts / groups, and treats
  // label-only rows ("FIXED COSTS") as group headings.
  function detectBudget(rows) {
    rows = (rows || []).map(function (r) { return (r || []).map(function (c) { return c == null ? "" : c; }); }).filter(function (r) { return r.some(function (c) { return String(c).trim() !== ""; }); });
    var isNum = function (v) { return typeof v === "number" || (/\d/.test(String(v)) && !isNaN(parseAmount(v)) && /^[\sR$€£ZA()\-\d.,]+$/i.test(String(v).trim())); };
    var hi = -1, cols = { name: -1, amt: -1, group: -1, kw: -1 };
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
      var r = rows[i].map(function (c) { return String(c).toLowerCase().trim(); });
      var n = r.findIndex(function (c) { return /^(category|categories|item|items|expense|expenses|name|description|budget item|line item|what|type of expense)$/.test(c) || /categ|expense item|description/.test(c); });
      var a = r.findIndex(function (c) { return /budget|amount|monthly|planned|plan|value|cost|zar|rand|^r$|per month|allocation|limit/.test(c) && !/actual|spent|difference|variance|categ/.test(c); });
      if (n > -1 || (a > -1 && r.filter(function (c) { return c; }).length >= 2)) {
        hi = i; cols.name = n; cols.amt = a;
        cols.group = r.findIndex(function (c, j) { return j !== n && /^(group|type|section|class|bucket|kind)$/.test(c); });
        cols.kw = r.findIndex(function (c) { return /keyword|merchant|match|shops?/.test(c); });
        break;
      }
    }
    var body = rows.slice(hi + 1);
    var width = Math.max.apply(null, rows.map(function (r) { return r.length; }).concat([0]));
    if (cols.amt < 0) { // first mostly-numeric column
      var bestC = -1, bestN = 0;
      for (var c = 0; c < width; c++) { var cnt = body.filter(function (r) { return isNum(r[c]) && parseAmount(r[c]) > 0; }).length; if (cnt > bestN) { bestN = cnt; bestC = c; } }
      cols.amt = bestC;
    }
    if (cols.name < 0) { // first mostly-text column
      for (var c2 = 0; c2 < width; c2++) { if (c2 === cols.amt) continue; var t = body.filter(function (r) { return String(r[c2]).trim() && !isNum(r[c2]); }).length; if (t >= Math.max(1, body.length * 0.4)) { cols.name = c2; break; } }
    }
    var items = [], income = 0, group = "";
    if (cols.name < 0 || cols.amt < 0) return { items: items, income: 0, cols: cols };
    body.forEach(function (r) {
      var name = String(r[cols.name] || "").trim(), raw = r[cols.amt], amt = Math.abs(parseAmount(raw));
      var anyNum = r.some(function (v, j) { return j !== cols.name && isNum(v) && parseAmount(v) !== 0; });
      if (name && !anyNum && name.length < 40) { group = titleCase(name.replace(/[:\-]+$/, "")); return; } // heading row
      if (!name || !(amt > 0)) return;
      if (/^(sub)?total|grand total|^sum$|balance|surplus|remaining|left over|difference/i.test(name)) return;
      if (/income|salary|wage|take.?home|net pay/i.test(name) && !/tax/i.test(name)) { income += amt; return; }
      items.push({
        name: name.slice(0, 40), budget: Math.round(amt * 100) / 100,
        group: cols.group > -1 && String(r[cols.group]).trim() ? titleCase(String(r[cols.group]).trim()) : (group || ""),
        keywords: cols.kw > -1 ? String(r[cols.kw] || "").trim() : "", on: true
      });
    });
    return { items: items, income: income, cols: cols };
  }

  function importRows(rows, fname) {
    var det = detectBudget(rows);
    if (!det.items.length) { toast("I couldn't find categories and amounts in that file. Try the template."); return; }
    var mode = cats().length ? "merge" : "replace";
    function draw() {
      var tot = det.items.filter(function (i) { return i.on; }).reduce(function (a, i) { return a + i.budget; }, 0);
      var html = sheetHead("Import budget", esc(fname) + " · " + det.items.length + " categories found") +
        (cats().length ? '<div class="seg" id="impMode" style="margin-bottom:12px"><button type="button" data-m="merge" class="' + (mode === "merge" ? "on" : "") + '">Add & update</button><button type="button" data-m="replace" class="' + (mode === "replace" ? "on" : "") + '">Replace all</button></div>' : "") +
        (det.income ? '<label class="check" style="margin-bottom:10px"><input type="checkbox" id="impIncome" checked> Use ' + money(det.income) + ' as my monthly income</label>' : "") +
        '<div style="max-height:46dvh;overflow:auto;margin:0 -4px"><table class="imp-table"><tbody>' + det.items.map(function (it, i) {
          return '<tr><td style="width:30px"><input type="checkbox" data-i="' + i + '" data-f="on" ' + (it.on ? "checked" : "") + ' style="width:20px;height:20px;accent-color:var(--accent)"></td><td style="width:34px;font-size:20px">' + emojiFor(it.name + " " + it.group) + '</td>' +
            '<td><input data-i="' + i + '" data-f="name" value="' + esc(it.name) + '"><div class="tiny faint" style="padding:2px 4px">' + esc(it.group || "Expenses") + '</div></td><td style="width:110px"><input data-i="' + i + '" data-f="budget" inputmode="decimal" value="' + it.budget + '" class="num"></td></tr>';
        }).join("") + '</tbody></table></div>' +
        '<div class="summary-line total"><span>Total monthly budget</span><span class="num">' + money(tot) + '</span></div>' +
        '<button class="btn primary block" id="impGo" type="button" style="margin-top:8px">Import ' + det.items.filter(function (i) { return i.on; }).length + ' categories</button>' +
        '<p class="small faint" style="text-align:center">' + (mode === "merge" ? "Existing categories with the same name get the new amount. Nothing else changes." : "Your current categories are replaced. Transactions stay; any in removed categories become unsorted.") + '</p>';
      openSheet(html, function (el) {
        $$("[data-m]", el).forEach(function (b) { b.onclick = function () { mode = b.dataset.m; draw(); }; });
        el.addEventListener("change", function (e) {
          var inp = e.target.closest("[data-i]"); if (!inp) return;
          var it = det.items[+inp.dataset.i];
          if (inp.dataset.f === "on") it.on = inp.checked; else if (inp.dataset.f === "budget") it.budget = Math.abs(parseAmount(inp.value)) || 0; else it.name = inp.value.trim();
          if (inp.dataset.f !== "name") draw();
        });
        $("#impGo", el).onclick = function () {
          var chosen = det.items.filter(function (i) { return i.on && i.name; });
          if (mode === "replace") {
            var keep = {};
            cats().forEach(function (c) { keep[c.name.toLowerCase()] = c; });
            var names = chosen.map(function (i) { return i.name.toLowerCase(); });
            cats().forEach(function (c) { if (names.indexOf(c.name.toLowerCase()) < 0) { c.deleted = true; touch(c); txs().forEach(function (t) { if (t.catId === c.id) { t.catId = null; t.needsReview = true; touch(t); } }); } });
          }
          var base = cats().length;
          chosen.forEach(function (it, i) {
            var ex = cats().filter(function (c) { return c.name.toLowerCase() === it.name.toLowerCase(); })[0];
            if (ex) { ex.budget = it.budget; if (it.group) ex.group = it.group; if (it.keywords) ex.keywords = it.keywords; touch(ex); }
            else db.categories.push(touch({ id: uid(), name: it.name, emoji: emojiFor(it.name + " " + it.group), budget: it.budget, group: it.group || "Expenses", keywords: it.keywords, kind: guessKind(it.name + " " + it.group), order: base + i }));
          });
          if (det.income && $("#impIncome", el) && $("#impIncome", el).checked) { db.settings.income = det.income; touch(db.settings); }
          local.onboarded = true; saveLocal();
          closeSheet(); ui.tab = "home"; commit(); toast("Imported " + chosen.length + " categories ✓");
        };
      });
    }
    draw();
  }

  function downloadFile(name, content, type) {
    var blob = new Blob([content], { type: type });
    var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function templateCsv() {
    return "Category,Monthly budget,Group,Keywords\n" + STARTER.map(function (s) { return [s[1], s[2], s[3], s[4] || ""].map(function (v) { return /[,"]/.test(v) ? '"' + String(v).replace(/"/g, '""') + '"' : v; }).join(","); }).join("\n") + "\n";
  }

  // ======================================================================
  // Starter categories & sample data
  // ======================================================================
  var STARTER = [
    ["🏠", "Rent / Bond", 9500, "Fixed costs", "rent, bond, levy"],
    ["💡", "Electricity & water", 1400, "Fixed costs", ""],
    ["📱", "Phone & internet", 899, "Fixed costs", ""],
    ["🛡️", "Insurance", 1250, "Fixed costs", ""],
    ["💊", "Medical aid & health", 2300, "Fixed costs", ""],
    ["📺", "Subscriptions", 450, "Fixed costs", ""],
    ["🛒", "Groceries", 4500, "Everyday", ""],
    ["⛽", "Fuel", 2200, "Everyday", ""],
    ["🚕", "Transport", 500, "Everyday", ""],
    ["🍔", "Eating out", 1200, "Lifestyle", ""],
    ["🛍️", "Shopping", 1000, "Lifestyle", ""],
    ["🎉", "Entertainment", 600, "Lifestyle", ""],
    ["🎁", "Gifts & giving", 400, "Lifestyle", ""],
    ["🏦", "Savings", 2500, "Goals", ""]
  ];
  function addStarter() {
    STARTER.forEach(function (s, i) { db.categories.push(touch({ id: uid(), emoji: s[0], name: s[1], budget: s[2], group: s[3], keywords: s[4], kind: guessKind(s[1]), order: i })); });
    if (!+db.settings.income) { db.settings.income = 32000; touch(db.settings); }
  }
  function loadSample() {
    addStarter();
    db.categories.forEach(function (c) { c.sample = true; });
    var shops = {
      "Groceries": [["Checkers Hyper", 380, 1100], ["Woolworths Food", 150, 650], ["Pick n Pay", 120, 700], ["Spar", 60, 260]],
      "Fuel": [["Engen", 600, 900], ["Shell", 500, 850]],
      "Eating out": [["Mugg & Bean", 120, 320], ["Uber Eats", 140, 290], ["Nando's", 180, 340]],
      "Transport": [["Uber", 60, 180], ["Bolt", 50, 140]],
      "Shopping": [["Takealot", 199, 900], ["Mr Price", 150, 520]],
      "Entertainment": [["Ster-Kinekor", 120, 260], ["Computicket", 180, 450]],
      "Subscriptions": [["Netflix", 199, 199], ["Spotify", 69.99, 69.99]],
      "Phone & internet": [["Vodacom", 899, 899]],
      "Electricity & water": [["City Power prepaid", 500, 800]]
    };
    var byName = {}; cats().forEach(function (c) { byName[c.name] = c; });
    var seed = 7; var rnd = function () { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; };
    for (var m = 5; m >= 0; m--) {
      var p = period(-m), lastDay = m === 0 ? p.elapsed : p.days;
      var add = function (date, amt, merch, c, type) { db.transactions.push(touch({ id: uid(), date: iso(date), amount: Math.round(amt * 100) / 100, type: type || "expense", merchant: merch, catId: c ? c.id : null, note: "", source: rnd() > 0.5 ? "sms" : "manual", sample: true })); };
      add(p.start, 32000, "Salary", null, "income");
      if (lastDay >= 1) { add(p.start, 9500, "Rent", byName["Rent / Bond"]); add(addDays(p.start, 0), 2500, "EasyEquities TFSA", byName["Savings"]); }
      if (lastDay >= 3) { add(addDays(p.start, 2), 1250, "Outsurance", byName["Insurance"]); add(addDays(p.start, 2), 2300, "Discovery Health", byName["Medical aid & health"]); }
      Object.keys(shops).forEach(function (cn) {
        var c = byName[cn]; var n = cn === "Groceries" ? 9 : cn === "Fuel" ? 3 : cn === "Eating out" ? 4 : 2;
        for (var k = 0; k < n; k++) {
          var d = Math.floor(rnd() * p.days); if (d >= lastDay) continue;
          var s = shops[cn][Math.floor(rnd() * shops[cn].length)];
          add(addDays(p.start, d), s[1] + rnd() * (s[2] - s[1]) * (m === 0 && cn === "Eating out" ? 1.6 : 1), s[0], c);
        }
      });
    }
    // a couple of SMS waiting to be sorted
    ingestSms("FNB :-) R389.00 paid from cheq a/c..123456 @ Builders Warehouse Rivonia. Avail R8 412.30. " + shortDate(today()) + " 10:14", null, { sample: true });
    ingestSms("Absa: CHEQ1234, Pur, " + slashDate(today()).replace(/\/(\d{2})(\d{2})$/, "/$2") + " MAKRO WOODMEAD, R-1,249.00, Available R7,163.30.", null, { sample: true });
    ingestSms("Capitec: Purchase of R72.50 at SAMS CORNER on " + slashDate(today()) + ". Avail bal R3 120.44", null, { sample: true });
    local.onboarded = true; saveLocal();
  }
  function shortDate(d) { return d.getDate() + MONTHS[d.getMonth()].slice(0, 3); }
  function slashDate(d) { return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear(); }

  // ======================================================================
  // Sync (end-to-end encrypted) via Cloudflare Worker
  // ======================================================================
  var syncTimer = null, syncing = false, syncState = "idle", cryptoKey = null, cryptoFor = "";
  function b64(buf) { var s = "", b = new Uint8Array(buf); for (var i = 0; i < b.length; i++) s += String.fromCharCode(b[i]); return btoa(s); }
  function unb64(s) { var bin = atob(s), b = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); return b; }
  function getKey() {
    if (cryptoKey && cryptoFor === local.syncCode) return Promise.resolve(cryptoKey);
    var enc = new TextEncoder();
    return crypto.subtle.importKey("raw", enc.encode(local.syncCode), "PBKDF2", false, ["deriveKey"]).then(function (base) {
      return crypto.subtle.deriveKey({ name: "PBKDF2", salt: enc.encode("budgie-sync-v1"), iterations: 150000, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
    }).then(function (k) { cryptoKey = k; cryptoFor = local.syncCode; return k; });
  }
  function encryptDoc(doc) {
    var iv = crypto.getRandomValues(new Uint8Array(12));
    return getKey().then(function (k) { return crypto.subtle.encrypt({ name: "AES-GCM", iv: iv }, k, new TextEncoder().encode(JSON.stringify(doc))); })
      .then(function (ct) { return { v: 1, iv: b64(iv), ct: b64(ct) }; });
  }
  function decryptDoc(box) {
    if (!box || !box.ct) return Promise.resolve(null);
    return getKey().then(function (k) { return crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(box.iv) }, k, unb64(box.ct)); })
      .then(function (pt) { return JSON.parse(new TextDecoder().decode(pt)); });
  }
  function mergeDocs(a, b) {
    var out = blankDb();
    out.settings = (b.settings && (b.settings.updatedAt || 0) > (a.settings.updatedAt || 0)) ? b.settings : a.settings;
    ["categories", "transactions", "rules"].forEach(function (k) {
      var map = {};
      (a[k] || []).concat(b[k] || []).forEach(function (r) { var ex = map[r.id]; if (!ex || (r.updatedAt || 0) > (ex.updatedAt || 0)) map[r.id] = r; });
      out[k] = Object.keys(map).map(function (id) { return map[id]; });
    });
    return out;
  }
  function api(path, body) {
    return fetch(local.syncUrl.replace(/\/+$/, "") + path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
      .then(function (r) { if (!r.ok) return r.text().then(function (t) { throw new Error(t || ("HTTP " + r.status)); }); return r.json(); });
  }
  function scheduleSync() { if (!local.syncCode) return; clearTimeout(syncTimer); syncTimer = setTimeout(syncNow, 2500); }
  function syncNow(manual) {
    if (!local.syncCode || !local.syncUrl || syncing) return Promise.resolve();
    syncing = true; syncState = "busy"; renderSyncState();
    var before = JSON.stringify(db), added = 0, acks = [];
    return api("/v1/pull", { code: local.syncCode }).then(function (res) {
      return decryptDoc(res.doc).then(function (remote) {
        var remoteStr = remote ? JSON.stringify(remote) : "";
        if (remote) db = mergeDocs(db, remote);
        (res.inbox || []).forEach(function (it) { var r = ingestSms(it.text, it.at); if (r.ok) added++; acks.push(it.id); });
        var merged = JSON.stringify(db);
        persist();
        if (merged !== remoteStr || acks.length) return encryptDoc(db).then(function (box) { return api("/v1/push", { code: local.syncCode, doc: box, ack: acks }); });
      });
    }).then(function () {
      local.lastSync = now(); saveLocal(); syncState = "ok";
      if (JSON.stringify(db) !== before) render();
      if (added) { var q = reviewQueue().length; toast("📩 " + added + " new from your bank" + (q ? " · " + q + " to sort" : ""), q ? "Sort" : null, q ? reviewSheet : null); }
      else if (manual) toast("Synced ✓");
    }).catch(function (e) {
      syncState = "err"; if (manual) toast("Sync failed: " + (e.message || e).toString().slice(0, 80));
      if (String(e.message).indexOf("OperationError") > -1) toast("Couldn't decrypt — is the sync code right?");
    }).then(function () { syncing = false; renderSyncState(); });
  }
  function renderSyncState() {
    var on = !!(local.syncCode && local.syncUrl);
    $("#syncBtn").hidden = !on;
    $("#syncDot").className = "sync-dot " + (syncState === "busy" ? "busy" : syncState === "err" ? "err" : syncState === "ok" ? "ok" : "");
  }
  function newCode() {
    var a = "abcdefghjkmnpqrstuvwxyz23456789", b = crypto.getRandomValues(new Uint8Array(20)), s = "";
    for (var i = 0; i < 20; i++) { s += a[b[i] % a.length]; if (i % 5 === 4 && i < 19) s += "-"; }
    return s;
  }
  function joinSheet(prefill) {
    var html = sheetHead("Join a sync space", "Use the code from your other device or your family’s phone") + '<div class="stack">' +
      '<label class="field">Sync server URL<input id="jUrl" value="' + esc((prefill && prefill.server) || local.syncUrl || DEFAULT_SYNC_URL) + '" placeholder="https://budgie-sync.yourname.workers.dev"></label>' +
      '<label class="field">Sync code<input id="jCode" value="' + esc((prefill && prefill.code) || "") + '" placeholder="xxxxx-xxxxx-xxxxx-xxxxx" autocapitalize="off" autocomplete="off"></label>' +
      '<p class="small faint" style="margin:0">Anything already on this device is merged in — nothing is lost.</p>' +
      '<button class="btn primary block" id="jGo" type="button">Join & sync</button></div>';
    openSheet(html, function (el) {
      $("#jGo", el).onclick = function () {
        var u = $("#jUrl", el).value.trim(), c = $("#jCode", el).value.trim().toLowerCase();
        if (!/^https:\/\/|^http:\/\/(localhost|127\.0\.0\.1)[:/]/.test(u)) { toast("The server URL should start with https://"); return; }
        if (c.replace(/-/g, "").length < 16) { toast("That code looks too short"); return; }
        local.syncUrl = u; local.syncCode = c; local.onboarded = true; saveLocal(); cryptoKey = null;
        closeSheet(); toast("Joining…");
        syncNow(true).then(function () { ui.tab = "home"; render(); });
      };
    });
  }

  function smsGuideSheet() {
    var ep = local.syncUrl.replace(/\/+$/, "") + "/v1/sms";
    var html = sheetHead("Forward bank SMS automatically", "Payments land in Budgie and get sorted for you") +
      '<div class="seg" id="osSeg" style="margin-bottom:14px"><button type="button" data-os="ios" class="on">iPhone</button><button type="button" data-os="android">Android</button></div>' +
      '<div class="small muted">Web address</div><div class="copy-box" style="margin-bottom:8px"><span>' + esc(ep) + '</span><button class="btn sm" data-copy="' + esc(ep) + '" type="button">Copy</button></div>' +
      '<div class="small muted">Your code</div><div class="copy-box" style="margin-bottom:14px"><span>' + esc(local.syncCode) + '</span><button class="btn sm" data-copy="' + esc(local.syncCode) + '" type="button">Copy</button></div>' +
      '<div id="g-ios"><ol class="steps">' +
      '<li>In your banking app, turn on <b>SMS notifications</b> for card purchases and payments.</li>' +
      '<li>Open the <b>Shortcuts</b> app → <b>Automation</b> tab → <b>+</b> → <b>Message</b>.</li>' +
      '<li>Set <b>Sender</b> to your bank’s SMS contact (e.g. the FNB or Capitec number) — or use <b>Message Contains</b> “R”. Choose <b>Run Immediately</b>, then <b>Next</b>.</li>' +
      '<li>Add the action <b>Get Contents of URL</b>. Paste the web address above.</li>' +
      '<li>Tap <b>Show More</b>: Method <b>POST</b>, Request Body <b>JSON</b>. Add a Text field <code>code</code> = your code, and a Text field <code>text</code> = the <b>Shortcut Input</b> variable (tap it and choose <i>Content</i>).</li>' +
      '<li>Tap <b>Done</b>. Send yourself a test by making a small purchase — it appears next time Budgie opens.</li></ol></div>' +
      '<div id="g-android" hidden><ol class="steps">' +
      '<li>In your banking app, turn on <b>SMS notifications</b> for transactions.</li>' +
      '<li>Install the free <b>MacroDroid</b> app from the Play Store.</li>' +
      '<li>Add Macro → Trigger: <b>SMS Received</b> → from your bank’s number (or any, with text containing “R”).</li>' +
      '<li>Action: <b>Web Interactions → HTTP Request</b>. Method <b>POST</b>, URL = the web address above.</li>' +
      '<li>Body: content type <code>application/json</code>, body <code>{"code":"' + esc(local.syncCode) + '","text":"[sms_message]"}</code></li>' +
      '<li>Save. Bank sends push notifications instead of SMS? Use the <b>Notification Received</b> trigger with <code>[notification]</code> instead.</li></ol></div>' +
      '<p class="small faint" style="margin-top:14px">Budgie only reads the amount, shop and date. Messages without an amount (like OTPs) are ignored. If it isn’t sure about a category, it asks you.</p>';
    openSheet(html, function (el) {
      $$("[data-os]", el).forEach(function (b) { b.onclick = function () { $$("[data-os]", el).forEach(function (x) { x.classList.toggle("on", x === b); }); $("#g-ios", el).hidden = b.dataset.os !== "ios"; $("#g-android", el).hidden = b.dataset.os !== "android"; }; });
    });
  }

  // ======================================================================
  // Backup
  // ======================================================================
  function restoreBackup(file) {
    var r = new FileReader();
    r.onload = function () {
      try {
        var d = JSON.parse(String(r.result));
        if (!d || !Array.isArray(d.categories)) throw new Error("Not a Budgie backup");
        db = mergeDocs(db, d); local.onboarded = true; saveLocal(); commit(); toast("Backup restored ✓");
      } catch (e) { toast("Couldn't restore: " + e.message); }
    };
    r.readAsText(file);
  }

  // ======================================================================
  // Toast, theme, install banner
  // ======================================================================
  var toastTimer;
  function toast(msg, actionLabel, action) {
    var el = $("#toast");
    el.innerHTML = "<span>" + esc(msg) + "</span>" + (actionLabel ? '<button type="button">' + esc(actionLabel) + "</button>" : "");
    el.hidden = false;
    if (actionLabel) $("button", el).onclick = function () { el.hidden = true; action(); };
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { el.hidden = true; }, actionLabel ? 6000 : 2800);
  }
  function applyTheme() {
    if (local.theme === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", local.theme);
    document.body.classList.toggle("simple", !!local.simple);
    if (local.simple && ui.tab === "insights") ui.tab = "home";
  }
  function isStandalone() { return window.navigator.standalone === true || window.matchMedia("(display-mode: standalone)").matches; }
  function isIOS() { return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1); }

  // ======================================================================
  // Events
  // ======================================================================
  function switchTab(tab) { ui.tab = tab; render(); $("#main").scrollTop = 0; }
  document.addEventListener("click", function (e) {
    var tabBtn = e.target.closest("[data-tab]");
    if (tabBtn && !e.target.closest("#sheet")) { switchTab(tabBtn.dataset.tab); return; }
    if (e.target.closest("[data-close]")) { closeSheet(); return; }
    var cp = e.target.closest("[data-copy]");
    if (cp) { var v = cp.dataset.copy; (navigator.clipboard ? navigator.clipboard.writeText(v) : Promise.reject()).then(function () { toast("Copied"); }, function () { window.prompt && window.prompt("Copy this:", v); }); return; }
    var th = e.target.closest("[data-theme-v]");
    if (th) { local.theme = th.dataset.themeV; saveLocal(); render(); return; }
    var a = e.target.closest("[data-act]"); if (!a) return;
    var id = a.dataset.id;
    switch (a.dataset.act) {
      case "cat": catDetailSheet(cat(id)); break;
      case "cat-edit": catSheet(cat(id)); break;
      case "new-cat": catSheet(null); break;
      case "tx": if (!e.target.closest("#sheet")) txSheet(db.transactions.filter(function (t) { return t.id === id; })[0]); break;
      case "review": reviewSheet(); break;
      case "filter": ui.actFilter = a.dataset.f; render(); break;
      case "pick-file": $("#fileInput").click(); break;
      case "template": downloadFile("budgie-budget-template.csv", templateCsv(), "text/csv"); break;
      case "starter": addStarter(); local.onboarded = true; saveLocal(); commit(); toast("Added suggested categories — tweak amounts on the Budget tab"); break;
      case "sample": loadSample(); commit(); break;
      case "clear-sample":
        db.transactions.forEach(function (t) { if (t.sample) { t.deleted = true; touch(t); } });
        db.categories.forEach(function (c) { if (c.sample) { c.deleted = true; touch(c); } });
        commit(); toast("Cleared — let’s set up your real budget"); break;
      case "join": joinSheet(); break;
      case "sync-now": syncNow(true); break;
      case "sms-guide": smsGuideSheet(); break;
      case "sync-create":
        var u = ($("#syncUrl") || {}).value || local.syncUrl;
        if (!/^https:\/\/|^http:\/\/(localhost|127\.0\.0\.1)[:/]/.test(u || "")) { toast("Enter your sync server URL first (https://…)"); return; }
        local.syncUrl = u.trim(); local.syncCode = newCode(); saveLocal(); cryptoKey = null;
        syncNow(true).then(function () { render(); smsGuideSheet(); }); break;
      case "sync-off":
        if (!a.dataset.armed) { a.dataset.armed = 1; a.textContent = "Tap again to confirm"; return; }
        local.syncCode = ""; saveLocal(); syncState = "idle"; render(); toast("Sync turned off on this device. Your data stays here."); break;
      case "export": downloadFile("budgie-backup-" + iso(today()) + ".json", JSON.stringify(db, null, 1), "application/json"); break;
      case "import-backup": $("#backupInput").click(); break;
      case "export-csv":
        downloadFile("budgie-transactions.csv", "Date,Type,Merchant,Category,Amount,Note\n" + txs().sort(txSort).map(function (t) { return [t.date, t.type, t.merchant, (cat(t.catId) || {}).name || "", t.amount, t.note || ""].map(function (v) { v = String(v); return /[,"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(","); }).join("\n"), "text/csv"); break;
      case "wipe":
        if (!a.dataset.armed) { a.dataset.armed = 1; a.textContent = "Tap again — this can’t be undone"; return; }
        localStorage.removeItem(STORE_KEY); db = blankDb(); local.onboarded = false; saveLocal(); ui.tab = "home"; render(); toast("Erased"); break;
      case "sms-preview": case "sms-add":
        var lines = ($("#smsPaste").value || "").split(/\n\s*\n|\n(?=\S*\s?(?:fnb|capitec|absa|std|standard|nedbank|discovery|tyme|investec))/i).map(function (s) { return s.trim(); }).filter(Boolean);
        if (lines.length <= 1) lines = ($("#smsPaste").value || "").split("\n").map(function (s) { return s.trim(); }).filter(Boolean);
        var out = $("#smsResult");
        if (a.dataset.act === "sms-preview") {
          out.innerHTML = lines.map(function (l) {
            var p = parseSms(l); if (p.error) return '<div class="card flat" style="padding:10px">⚠️ ' + esc(p.error) + '</div>';
            var fake = { merchant: p.merchant, raw: l, type: p.type }; var c = categorise(fake);
            return '<div class="card flat" style="padding:10px"><b>' + esc(p.merchant) + '</b> · ' + money(p.amount, true) + ' · ' + p.type + ' · ' + (p.date || "today") + (p.bank ? " · " + esc(p.bank) : "") +
              '<br><span class="faint">' + (p.type !== "expense" ? "No category needed" : c.confidence >= 0.8 ? "→ " + esc((cat(c.catId) || { name: "Transfer" }).name) + " (" + esc(c.reason) + ")" : "→ will ask you (" + esc(c.reason) + ")") + '</span></div>';
          }).join("") || '<span class="faint">Paste something first.</span>';
        } else {
          var ok = 0, msgs = [];
          lines.forEach(function (l) { var r = ingestSms(l, null); if (r.ok) ok++; else msgs.push(r.reason); });
          commit();
          toast(ok ? "Added " + ok + " transaction" + (ok > 1 ? "s" : "") : (msgs[0] || "Nothing added"));
          if (reviewQueue().length && ok) setTimeout(reviewSheet, 400);
        }
        break;
    }
  });
  document.addEventListener("input", function (e) {
    if (e.target.id === "actSearch") {
      ui.actQuery = e.target.value; var pos = e.target.selectionStart; render();
      var s = $("#actSearch"); s.focus(); try { s.setSelectionRange(pos, pos); } catch (er) {}
    }
  });
  document.addEventListener("change", function (e) {
    var t = e.target;
    if (t.id === "incomeInput") { db.settings.income = Math.abs(parseAmount(t.value)) || 0; touch(db.settings); commit(); }
    if (t.id === "startDay") { db.settings.startDay = +t.value; touch(db.settings); ui.offset = 0; commit(); }
    if (t.id === "currency") { db.settings.currency = t.value; touch(db.settings); commit(); }
    if (t.id === "simpleToggle") { local.simple = t.checked; saveLocal(); render(); }
  });
  $("#fileInput").addEventListener("change", function (e) { handleFile(e.target.files[0]); e.target.value = ""; });
  $("#backupInput").addEventListener("change", function (e) { if (e.target.files[0]) restoreBackup(e.target.files[0]); e.target.value = ""; });
  $("#addBtn").addEventListener("click", function () { if (!cats().length) { toast("Set up your budget first"); return; } txSheet(null); });
  $("#prevMonth").addEventListener("click", function () { ui.offset--; render(); });
  $("#nextMonth").addEventListener("click", function () { if (ui.offset < 0) { ui.offset++; render(); } else toast("That’s the current month"); });
  $("#syncBtn").addEventListener("click", function () { syncNow(true); });
  document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !$("#sheetWrap").hidden) closeSheet(); });

  // drag & drop a spreadsheet anywhere
  var dragDepth = 0;
  function hasFiles(e) { return e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], "Files") > -1; }
  window.addEventListener("dragenter", function (e) { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; $("#dropOverlay").hidden = false; var z = $("#dropZone"); if (z) z.classList.add("over"); });
  window.addEventListener("dragover", function (e) { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener("dragleave", function () { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) { $("#dropOverlay").hidden = true; var z = $("#dropZone"); if (z) z.classList.remove("over"); } });
  window.addEventListener("drop", function (e) {
    if (!hasFiles(e)) return; e.preventDefault(); dragDepth = 0; $("#dropOverlay").hidden = true;
    var z = $("#dropZone"); if (z) z.classList.remove("over");
    handleFile(e.dataTransfer.files[0]);
  });
  function mountDrop() { /* the drop zone is visual; the window-level handlers above do the work */ }

  // Join links: #join=CODE&server=URL (kept in the hash so it never hits a server log)
  function checkHash() {
    var h = location.hash.slice(1); if (!h) return;
    var params = {}; h.split("&").forEach(function (kv) { var p = kv.split("="); params[p[0]] = decodeURIComponent(p[1] || ""); });
    if (params.join) { history.replaceState(null, "", location.pathname); joinSheet({ code: params.join, server: params.server }); }
  }

  // install banner
  (function () {
    var dismissed = false;
    try { dismissed = localStorage.getItem("budgie_install_dismissed") === "1"; } catch (e) {}
    if (isIOS() && !isStandalone() && !dismissed) $("#installBanner").hidden = false;
    $("#installClose").addEventListener("click", function () { $("#installBanner").hidden = true; try { localStorage.setItem("budgie_install_dismissed", "1"); } catch (e) {} });
  })();

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", function () { navigator.serviceWorker.register("sw.js").catch(function () {}); });
  }
  document.addEventListener("visibilitychange", function () { if (document.visibilityState === "visible") { syncNow(); render(); } });
  var rzT; window.addEventListener("resize", function () { clearTimeout(rzT); rzT = setTimeout(function () { if (ui.tab === "insights") mountCharts(period(ui.offset)); }, 200); });
  setInterval(function () { if (document.visibilityState === "visible") syncNow(); }, 60000);

  render();
  checkHash();
  syncNow();

  // exposed for testing in the console
  window.Budgie = { handleFile: handleFile, parseSms: parseSms, detectBudget: detectBudget, categorise: categorise, merchantKey: merchantKey, get db() { return db; } };
})();
