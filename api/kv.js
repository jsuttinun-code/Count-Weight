// Vercel Serverless Function: ที่เก็บข้อมูลกลางสำหรับเว็บ Count-Weight (ใช้ Upstash Redis จาก Vercel Marketplace)
const U = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const T = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;

async function rd(cmds) {
  const r = await fetch(U + "/pipeline", { method: "POST", headers: { Authorization: "Bearer " + T }, body: JSON.stringify(cmds) });
  const j = await r.json();
  if (!r.ok) throw new Error(JSON.stringify(j));
  return j.map((x) => { if (x.error) throw new Error(x.error); return x.result; });
}

module.exports = async (req, res) => {
  try {
    if (!U || !T) return res.status(500).json({ error: "ยังไม่ได้เชื่อม Upstash Redis กับโปรเจกต์ Vercel" });
    if (req.method === "POST") {
      const b = typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {};
      const rows = (b.rows || []).filter((x) => /^[a-z]+$/.test(x.coll || "") && x.id != null);
      if (!rows.length) return res.json({ ok: 1 });
      const last = (await rd([["INCRBY", "sc:seq", rows.length]]))[0];
      const cmds = [];
      rows.forEach((x, i) => {
        cmds.push(["HSET", "sc:h:" + x.coll, String(x.id), JSON.stringify({ d: x.data, x: x.deleted ? 1 : 0 })]);
        cmds.push(["ZADD", "sc:z:" + x.coll, last - rows.length + 1 + i, String(x.id)]);
      });
      await rd(cmds);
      return res.json({ ok: 1 });
    }
    const c = String(req.query.coll || "").replace(/[^a-z]/g, "");
    const since = Number(req.query.since || 0);
    const flat = (await rd([["ZRANGEBYSCORE", "sc:z:" + c, "(" + since, "+inf", "WITHSCORES", "LIMIT", 0, 500]]))[0];
    const items = [];
    for (let i = 0; i < flat.length; i += 2) items.push([flat[i], Number(flat[i + 1])]);
    if (!items.length) return res.json({ rows: [], next: since, more: false });
    const vals = (await rd([["HMGET", "sc:h:" + c, ...items.map((x) => x[0])]]))[0];
    res.json({
      rows: items.map((x, i) => { const v = vals[i] ? JSON.parse(vals[i]) : {}; return { id: x[0], data: v.d, deleted: !!v.x, seq: x[1] }; }),
      next: items[items.length - 1][1],
      more: items.length === 500,
    });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
};
