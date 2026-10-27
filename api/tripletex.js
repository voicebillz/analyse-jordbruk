/* Vercel-funksjon: hentar data frå Tripletex sitt API på vegner av brukaren.
   Nøklane kjem frå nettlesaren i kvar førespurnad og blir ikkje lagra her.
   Berre lesing (GET) mot eit fast utval endepunkt er tillate. */
const BASE = {prod: "https://tripletex.no/v2", test: "https://api-test.tripletex.tech/v2"};
const ALLOW = [/^\/ledger\/posting$/, /^\/balanceSheet$/, /^\/company\/>withLoginAccess$/, /^\/token\/session\/>whoAmI$/, /^\/company\/\d+$/];
const MAX_PAGES = 40, PAGE = 10000;

const pick = (o, path) => path.split(".").reduce((x, k) => (x == null ? x : x[k]), o);

async function readErr(r) {
  let t = ""; try { t = await r.text(); } catch (e) { /* ignorer */ }
  try { const j = JSON.parse(t); return j.message || (j.validationMessages || []).map(v => v.message).join("; ") || t.slice(0, 300); } catch (e) { return t.slice(0, 300); }
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") return res.status(405).json({error: "Bruk POST"});
  let body = req.body || {};
  if (typeof body === "string") { try { body = JSON.parse(body); } catch (e) { body = {}; } }
  const {consumerToken, employeeToken, env = "prod", companyId = 0, requests = []} = body;
  if (!consumerToken || !employeeToken) return res.status(400).json({error: "Mangler consumer token eller employee token"});
  if (!Array.isArray(requests) || !requests.length || requests.length > 8) return res.status(400).json({error: "Ugyldig førespurnad"});
  const base = BASE[env] || BASE.prod;

  const exp = new Date(Date.now() + 2 * 864e5).toISOString().slice(0, 10);
  const tq = new URLSearchParams({consumerToken, employeeToken, expirationDate: exp});
  let token;
  try {
    const t = await fetch(`${base}/token/session/:create?${tq}`, {method: "PUT", headers: {Accept: "application/json"}});
    if (!t.ok) return res.status(401).json({error: "Tripletex godtok ikke nøklene", status: t.status, detail: await readErr(t)});
    token = (await t.json()).value.token;
  } catch (e) { return res.status(502).json({error: "Fikk ikke kontakt med Tripletex", detail: String(e.message || e)}); }
  const auth = "Basic " + Buffer.from(`${Number(companyId) || 0}:${token}`).toString("base64");

  const results = [];
  for (const r of requests) {
    if (!r || typeof r.path !== "string" || !ALLOW.some(re => re.test(r.path))) { results.push({error: "Endepunktet er ikke tillatt"}); continue; }
    const q0 = Object.assign({}, r.query || {}), rows = [];
    let from = 0, done = false;
    for (let page = 0; page < MAX_PAGES && !done; page++) {
      const q = new URLSearchParams(Object.assign({}, q0, r.paged === false ? {} : {from: String(from), count: String(PAGE)}));
      const x = await fetch(`${base}${r.path}?${q}`, {headers: {Authorization: auth, Accept: "application/json"}});
      if (!x.ok) { results.push({error: "Tripletex svarte med feil", status: x.status, detail: await readErr(x)}); done = true; break; }
      const j = await x.json();
      if (!Array.isArray(j.values)) { results.push({value: Array.isArray(r.pick) ? r.pick.map(p => pick(j.value, p)) : j.value}); done = true; break; }
      j.values.forEach(v => rows.push(Array.isArray(r.pick) ? r.pick.map(p => pick(v, p)) : v));
      from += j.values.length;
      if (r.paged === false || !j.values.length || from >= (j.fullResultSize || 0)) { results.push({values: rows, full: j.fullResultSize}); done = true; }
    }
    if (!done) results.push({values: rows, truncated: true});
  }
  fetch(`${base}/token/session/${encodeURIComponent(token)}`, {method: "DELETE", headers: {Authorization: auth}}).catch(() => {});
  return res.status(200).json({results});
};
