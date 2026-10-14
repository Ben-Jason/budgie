// Budgie sync Worker.
//
// Stores ONE encrypted blob per sync space (the app encrypts it on-device
// with a key derived from the sync code, so this Worker never sees your
// budget), plus a short-lived inbox of bank SMS texts posted by an iOS
// Shortcut / Android automation. The app pulls the inbox, parses the SMS
// locally, and acks them so they're deleted here.
//
// KV keys are derived from SHA-256 of the sync code, never the code itself.
//
//   POST /v1/pull  {code}               -> {doc, inbox:[{id,text,at}]}
//   POST /v1/push  {code, doc, ack:[]}  -> {ok}
//   POST /v1/sms   {code, text}         -> {ok}     (also accepts ?code= with a text/plain body)

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};
const MAX_DOC = 4 * 1024 * 1024;   // 4 MB encrypted doc
const MAX_SMS = 1500;              // chars per message
const MAX_INBOX = 300;             // pending messages per space
const SMS_TTL = 60 * 60 * 24 * 45; // inbox items expire after 45 days if never pulled

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "Content-Type": "application/json", ...CORS } });

async function spaceId(code) {
  const c = String(code || "").trim().toLowerCase();
  if (!/^[a-z0-9-]{16,64}$/.test(c) || c.replace(/-/g, "").length < 16) return null;
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("budgie:" + c));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 40);
}

async function readBody(request, url) {
  const type = request.headers.get("content-type") || "";
  const raw = await request.text();
  if (raw.length > MAX_DOC + 1024) throw new Error("too large");
  if (type.includes("application/json") || /^\s*\{/.test(raw)) {
    try { return JSON.parse(raw); } catch { /* fall through */ }
  }
  if (type.includes("application/x-www-form-urlencoded")) return Object.fromEntries(new URLSearchParams(raw));
  return { code: url.searchParams.get("code"), text: raw };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { headers: CORS });
    if (request.method === "GET" && url.pathname === "/") return json({ ok: true, service: "budgie-sync" });
    if (request.method !== "POST") return json({ error: "not found" }, 404);

    let body;
    try { body = await readBody(request, url); } catch { return json({ error: "bad request" }, 400); }
    const id = await spaceId(body.code || url.searchParams.get("code"));
    if (!id) return json({ error: "invalid sync code" }, 401);

    if (url.pathname === "/v1/sms") {
      const text = String(body.text || body.message || body.sms || "").trim().slice(0, MAX_SMS);
      if (!text) return json({ error: "no text" }, 400);
      const pending = await env.DATA.list({ prefix: `sms:${id}:`, limit: MAX_INBOX });
      if (pending.keys.length >= MAX_INBOX) return json({ error: "inbox full — open Budgie to sync" }, 429);
      const key = `sms:${id}:${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      await env.DATA.put(key, JSON.stringify({ text, at: Date.now() }), { expirationTtl: SMS_TTL });
      return json({ ok: true });
    }

    if (url.pathname === "/v1/pull") {
      const doc = await env.DATA.get(`doc:${id}`, "json");
      const list = await env.DATA.list({ prefix: `sms:${id}:`, limit: MAX_INBOX });
      const inbox = (await Promise.all(list.keys.map(async (k) => {
        const v = await env.DATA.get(k.name, "json");
        return v ? { id: k.name.slice(`sms:${id}:`.length), text: v.text, at: v.at } : null;
      }))).filter(Boolean);
      return json({ doc, inbox });
    }

    if (url.pathname === "/v1/push") {
      const doc = body.doc;
      if (!doc || typeof doc.ct !== "string" || typeof doc.iv !== "string") return json({ error: "doc must be encrypted" }, 400);
      if (doc.ct.length > MAX_DOC) return json({ error: "doc too large" }, 413);
      await env.DATA.put(`doc:${id}`, JSON.stringify({ v: 1, iv: doc.iv, ct: doc.ct, at: Date.now() }));
      const ack = Array.isArray(body.ack) ? body.ack.slice(0, MAX_INBOX) : [];
      await Promise.all(ack.filter((a) => /^[\w-]{1,40}$/.test(a)).map((a) => env.DATA.delete(`sms:${id}:${a}`)));
      return json({ ok: true });
    }

    return json({ error: "not found" }, 404);
  },
};
