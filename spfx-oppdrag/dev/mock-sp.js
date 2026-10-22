// Simulert SharePoint REST for lokal testing av lagringen (bare det appen bruker).
export function createMockSharePoint(opts = {}) {
  const db = JSON.parse(localStorage.getItem('mock-sp') || 'null') || { list: null, fields: [], items: [], next: 1 };
  const save = () => localStorage.setItem('mock-sp', JSON.stringify(db));
  const log = [];
  const res = (status, body) => ({ ok: status >= 200 && status < 300, status, json: async () => body, text: async () => JSON.stringify(body || '') });
  const user = opts.user || 'Test Bruker';
  const item = i => ({ Id: i.Id, Title: i.Title, Oppgavetype: i.Oppgavetype, Periode: i.Periode, Modified: i.Modified, Editor: { Title: i.Editor }, '@odata.etag': `"${i.v}"` });
  const http = {
    log, db,
    async get(url) {
      log.push(['GET', url]);
      if (/GetStorageEntity/.test(url)) return res(200, { Value: null });
      if (/GetByTitle\('Oppdrag'\)\/items/.test(url)) {
        if (!db.list) return res(404);
        return res(200, { value: db.items.slice().sort((a, b) => b.Modified.localeCompare(a.Modified)).map(item) });
      }
      if (/GetByTitle\('Oppdrag'\)/.test(url)) return db.list ? res(200, { Id: 'x' }) : res(404);
      return res(404);
    },
    async post(url, body, headers = {}) {
      log.push(['POST', url, headers['X-HTTP-Method'] || '', body]);
      if (opts.readOnly) return res(403);
      const m = headers['X-HTTP-Method'];
      if (/_api\/web\/lists$/.test(url)) { db.list = body; save(); return res(201, {}); }
      if (/CreateFieldAsXml/.test(url)) { db.fields.push(body.parameters.SchemaXml); save(); return res(201, {}); }
      if (/GetByInternalNameOrTitle/.test(url)) return res(204);
      const im = /items\((\d+)\)/.exec(url);
      if (im) {
        const i = db.items.find(x => x.Id === +im[1]);
        if (!i) return res(404);
        const want = headers['IF-MATCH'];
        if (want && want !== '*' && want !== `"${i.v}"`) return res(412);
        if (m === 'DELETE') db.items = db.items.filter(x => x !== i);
        else Object.assign(i, body, { v: i.v + 1, Modified: new Date().toISOString(), Editor: user });
        save(); return res(204);
      }
      if (/items$/.test(url)) {
        db.items.push({ Id: db.next++, v: 1, Modified: new Date().toISOString(), Editor: user, ...body });
        save(); return res(201, {});
      }
      return res(400);
    }
  };
  return http;
}
