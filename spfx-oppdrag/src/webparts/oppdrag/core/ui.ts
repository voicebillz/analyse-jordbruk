/**
 * Brukergrensesnitt for oppdragslisten. Rein DOM-kode uten SPFx-avhengigheter,
 * slik at den kan testes i en vanlig nettleser mot en simulert lagring.
 */
import { ConflictError, IOppdrag, IOppdragInput, IOppdragStore, OPPGAVETYPER } from './store';

export interface IAppOptions {
  dark?: boolean;
  refreshSeconds?: number;
}

export interface IAppHandle {
  setDark(dark: boolean): void;
  refresh(): Promise<void>;
  dispose(): void;
}

const MND = ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Desember'];

/** Forslag til periode ut fra oppgavetype og dagens dato. */
export function periodeForslag(type: string, now: Date = new Date()): string[] {
  const y = now.getFullYear();
  const m = now.getMonth();
  const out: string[] = [];
  if (type === 'MVA') {
    const t = Math.floor(m / 2) + 1;
    for (let i = 0; i < 6; i++) {
      let tt = t - i;
      let yy = y;
      while (tt < 1) { tt += 6; yy--; }
      out.push(`${tt}. termin ${yy}`);
    }
    out.push(`År ${y - 1}`);
  } else if (type === 'Årsoppgjør' || type === 'Skattemelding') {
    out.push(String(y - 1), String(y), String(y - 2));
  } else {
    for (let i = 0; i < 12; i++) {
      const d = new Date(y, m - i, 1);
      out.push(`${MND[d.getMonth()]} ${d.getFullYear()}`);
    }
    out.push(String(y), String(y - 1));
  }
  return out;
}

const CSS = `
.opd{--bg:#fff;--fg:#1f2328;--mut:#5b6470;--line:#e3e6ea;--soft:#f5f7f9;--acc:#2f6f4e;--accfg:#fff;--danger:#b42318;--chip:#eef4f0;--chipfg:#2f6f4e;--shadow:0 12px 40px rgba(0,0,0,.18);
  font-family:"Segoe UI",system-ui,-apple-system,sans-serif;font-size:14px;color:var(--fg);background:var(--bg);padding:20px 24px;box-sizing:border-box;min-height:100%}
.opd.dark{--bg:#1b1d21;--fg:#e8eaed;--mut:#a0a7b1;--line:#30343a;--soft:#24272c;--acc:#6fbf8e;--accfg:#10261a;--danger:#ff8a80;--chip:#20352a;--chipfg:#9fe0b7;--shadow:0 12px 40px rgba(0,0,0,.5)}
.opd *{box-sizing:border-box}
.opd h1{font-size:22px;font-weight:600;margin:0}
.opd .top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.opd .top .cnt{color:var(--mut)}
.opd .sp{flex:1}
.opd .bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}
.opd input,.opd select{font:inherit;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:8px 10px;min-width:0}
.opd input:focus,.opd select:focus{outline:2px solid var(--acc);outline-offset:-1px}
.opd .bar .q{flex:1 1 220px}
.opd button{font:inherit;cursor:pointer;border-radius:8px;padding:8px 14px;border:1px solid var(--line);background:var(--soft);color:var(--fg)}
.opd button.pri{background:var(--acc);border-color:var(--acc);color:var(--accfg);font-weight:600}
.opd button.del{color:var(--danger)}
.opd button.lnk{background:none;border:none;padding:4px 8px;color:var(--mut)}
.opd button.lnk:hover{color:var(--fg);background:var(--soft)}
.opd button:disabled{opacity:.5;cursor:default}
.opd table{width:100%;border-collapse:collapse}
.opd th{text-align:left;font-weight:600;color:var(--mut);font-size:12px;text-transform:uppercase;letter-spacing:.04em;padding:8px 10px;border-bottom:1px solid var(--line);cursor:pointer;user-select:none;white-space:nowrap}
.opd th.na{cursor:default}
.opd td{padding:10px;border-bottom:1px solid var(--line);vertical-align:middle}
.opd tr:hover td{background:var(--soft)}
.opd td.k{font-weight:600}
.opd td.m{color:var(--mut);font-size:13px}
.opd td.a{text-align:right;white-space:nowrap}
.opd .chip{display:inline-block;padding:2px 10px;border-radius:999px;background:var(--chip);color:var(--chipfg);font-size:12px;font-weight:600}
.opd .empty{padding:40px 10px;text-align:center;color:var(--mut)}
.opd .msg{padding:10px 12px;border-radius:8px;margin-bottom:12px;background:var(--soft);border:1px solid var(--line)}
.opd .msg.err{border-color:var(--danger);color:var(--danger)}
.opd .ov{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:flex-start;justify-content:center;padding:10vh 16px;z-index:1000}
.opd .dlg{background:var(--bg);color:var(--fg);border-radius:12px;box-shadow:var(--shadow);width:100%;max-width:440px;padding:20px}
.opd .dlg h2{font-size:18px;margin:0 0 14px}
.opd .dlg label{display:block;font-size:12px;font-weight:600;color:var(--mut);margin:12px 0 4px}
.opd .dlg input,.opd .dlg select{width:100%}
.opd .dlg .act{display:flex;gap:8px;justify-content:flex-end;margin-top:18px}
.opd .dlg .err{color:var(--danger);margin-top:10px;min-height:1em}
.opd .sugg{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}
.opd .sugg button{padding:3px 10px;border-radius:999px;font-size:12px}
@media (max-width:640px){.opd{padding:14px}.opd td.m.e,.opd th.e{display:none}.opd td,.opd th{padding:8px 4px}.opd td.a button{display:block;margin-left:auto;padding:3px 4px}}
`;

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function fmtDato(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const p = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

type SortKey = 'klient' | 'oppgavetype' | 'periode' | 'endret';

export function mountOppdragApp(host: HTMLElement, store: IOppdragStore, opts: IAppOptions = {}): IAppHandle {
  let items: IOppdrag[] = [];
  let loaded = false;
  let error = '';
  let q = '';
  let fType = '';
  let sort: SortKey = 'endret';
  let asc = false;
  let disposed = false;

  if (!document.getElementById('opd-css')) {
    const st = document.createElement('style');
    st.id = 'opd-css';
    st.textContent = CSS;
    document.head.appendChild(st);
  }
  const root = document.createElement('div');
  root.className = 'opd' + (opts.dark ? ' dark' : '');
  host.innerHTML = '';
  host.appendChild(root);

  root.innerHTML = `
    <div class="top"><h1>Bokføringsoppdrag</h1><span class="cnt"></span><span class="sp"></span>
      <button class="pri" data-a="ny">+ Nytt oppdrag</button></div>
    <div class="bar">
      <input class="q" type="search" placeholder="Søk etter klient eller periode">
      <select class="ft"><option value="">Alle oppgavetyper</option>${OPPGAVETYPER.map(t => `<option>${esc(t)}</option>`).join('')}</select>
    </div>
    <div class="msgs"></div>
    <div class="list"></div>`;

  const $ = <T extends HTMLElement>(sel: string): T => root.querySelector(sel) as T;
  const qInput = $<HTMLInputElement>('.q');
  const ftSel = $<HTMLSelectElement>('.ft');

  function visible(): IOppdrag[] {
    const s = q.trim().toLowerCase();
    const r = items.filter(o =>
      (!fType || o.oppgavetype === fType) &&
      (!s || `${o.klient} ${o.periode} ${o.oppgavetype}`.toLowerCase().indexOf(s) >= 0));
    r.sort((a, b) => {
      const c = String(a[sort]).localeCompare(String(b[sort]), 'nb', { numeric: true, sensitivity: 'base' });
      return asc ? c : -c;
    });
    return r;
  }

  function render(): void {
    const v = visible();
    $('.cnt').textContent = loaded ? `${items.length} oppdrag` : '';
    $('.msgs').innerHTML = error ? `<div class="msg err">${esc(error)}</div>` : '';
    const list = $('.list');
    if (!loaded) {
      list.innerHTML = '<div class="empty">Henter oppdrag …</div>';
      return;
    }
    if (!v.length) {
      list.innerHTML = `<div class="empty">${items.length ? 'Ingen oppdrag passer søket.' : 'Ingen oppdrag ennå. Trykk «+ Nytt oppdrag» for å legge til det første.'}</div>`;
      return;
    }
    const th = (k: SortKey, t: string, cls = ''): string =>
      `<th data-s="${k}" class="${cls}">${t}${sort === k ? (asc ? ' ▲' : ' ▼') : ''}</th>`;
    list.innerHTML = `<table><thead><tr>${th('klient', 'Klient')}${th('oppgavetype', 'Oppgavetype')}${th('periode', 'Periode')}${th('endret', 'Sist endret', 'e')}<th class="na"></th></tr></thead><tbody>${
      v.map(o => `<tr data-id="${o.id}">
        <td class="k">${esc(o.klient)}</td>
        <td>${o.oppgavetype ? `<span class="chip">${esc(o.oppgavetype)}</span>` : ''}</td>
        <td>${esc(o.periode)}</td>
        <td class="m e">${esc(fmtDato(o.endret))}${o.endretAv ? ` · ${esc(o.endretAv)}` : ''}</td>
        <td class="a"><button class="lnk" data-a="endre">Endre</button><button class="lnk del" data-a="slett">Slett</button></td>
      </tr>`).join('')}</tbody></table>`;
  }

  async function refresh(): Promise<void> {
    try {
      items = await store.list();
      error = '';
    } catch (e) {
      error = (e as Error).message;
    }
    loaded = true;
    if (!disposed) render();
  }

  const closers = new WeakMap<HTMLElement, () => void>();
  function overlay(html: string): HTMLElement {
    const ov = document.createElement('div');
    ov.className = 'ov';
    ov.innerHTML = `<div class="dlg" role="dialog" aria-modal="true">${html}</div>`;
    root.appendChild(ov);
    // eslint-disable-next-line prefer-const
    let close: () => void;
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') close(); };
    close = (): void => { ov.remove(); document.removeEventListener('keydown', onKey); };
    document.addEventListener('keydown', onKey);
    ov.addEventListener('mousedown', e => { if (e.target === ov) close(); });
    closers.set(ov, close);
    return ov;
  }
  const closeOv = (ov: HTMLElement): void => { const c = closers.get(ov); if (c) c(); };

  function openEdit(o?: IOppdrag): void {
    const klienter = Array.from(new Set(items.map(i => i.klient).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'nb'));
    const type = o ? o.oppgavetype : (fType || OPPGAVETYPER[0]);
    const typer = OPPGAVETYPER.indexOf(type) < 0 && type ? OPPGAVETYPER.concat([type]) : OPPGAVETYPER;
    const ov = overlay(`
      <h2>${o ? 'Endre oppdrag' : 'Nytt oppdrag'}</h2>
      <label for="opd-k">Klient</label>
      <input id="opd-k" list="opd-kl" autocomplete="off" value="${esc(o ? o.klient : '')}">
      <datalist id="opd-kl">${klienter.map(k => `<option value="${esc(k)}">`).join('')}</datalist>
      <label for="opd-t">Oppgavetype</label>
      <select id="opd-t">${typer.map(t => `<option${t === type ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>
      <label for="opd-p">Periode</label>
      <input id="opd-p" autocomplete="off" value="${esc(o ? o.periode : '')}">
      <div class="sugg"></div>
      <div class="err"></div>
      <div class="act"><button data-a="avbryt">Avbryt</button><button class="pri" data-a="lagre">Lagre</button></div>`);
    const k = ov.querySelector('#opd-k') as HTMLInputElement;
    const t = ov.querySelector('#opd-t') as HTMLSelectElement;
    const p = ov.querySelector('#opd-p') as HTMLInputElement;
    const sugg = ov.querySelector('.sugg') as HTMLElement;
    const err = ov.querySelector('.err') as HTMLElement;
    const save = ov.querySelector('[data-a="lagre"]') as HTMLButtonElement;
    const fillSugg = (): void => {
      sugg.innerHTML = periodeForslag(t.value).slice(0, 6).map(s => `<button type="button">${esc(s)}</button>`).join('');
    };
    fillSugg();
    if (!o && !p.value) p.value = periodeForslag(t.value)[0];
    t.addEventListener('change', () => { fillSugg(); if (!o) p.value = periodeForslag(t.value)[0]; });
    sugg.addEventListener('click', e => {
      const b = (e.target as HTMLElement).closest('button');
      if (b) p.value = b.textContent || '';
    });
    const submit = async (): Promise<void> => {
      const data: IOppdragInput = { klient: k.value.trim(), oppgavetype: t.value, periode: p.value.trim() };
      if (!data.klient) { err.textContent = 'Skriv inn klient.'; k.focus(); return; }
      save.disabled = true;
      err.textContent = '';
      try {
        if (o) await store.update({ ...data, id: o.id, etag: o.etag });
        else await store.create(data);
        closeOv(ov);
        await refresh();
      } catch (e) {
        save.disabled = false;
        if (e instanceof ConflictError) {
          err.textContent = 'Noen andre har endret dette oppdraget i mellomtiden. Listen er oppdatert – åpne oppdraget på nytt.';
          await refresh();
        } else {
          err.textContent = (e as Error).message;
        }
      }
    };
    save.addEventListener('click', () => { submit().catch(() => undefined); });
    ov.addEventListener('keydown', e => { if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') submit().catch(() => undefined); });
    ov.querySelector('[data-a="avbryt"]')!.addEventListener('click', () => closeOv(ov));
    k.focus();
  }

  function openDelete(o: IOppdrag): void {
    const ov = overlay(`
      <h2>Slette oppdrag?</h2>
      <p>${esc(o.klient)} – ${esc(o.oppgavetype)} ${esc(o.periode)} blir slettet for alle.</p>
      <div class="err"></div>
      <div class="act"><button data-a="avbryt">Avbryt</button><button class="pri del" data-a="slett">Slett</button></div>`);
    const b = ov.querySelector('[data-a="slett"]') as HTMLButtonElement;
    b.addEventListener('click', () => {
      b.disabled = true;
      store.remove(o.id).then(() => { closeOv(ov); return refresh(); }).catch((e: Error) => {
        b.disabled = false;
        (ov.querySelector('.err') as HTMLElement).textContent = e.message;
      });
    });
    ov.querySelector('[data-a="avbryt"]')!.addEventListener('click', () => closeOv(ov));
  }

  root.addEventListener('click', e => {
    const el = e.target as HTMLElement;
    const th = el.closest('th[data-s]') as HTMLElement | null;
    if (th) {
      const s = th.getAttribute('data-s') as SortKey;
      if (sort === s) asc = !asc;
      else { sort = s; asc = s !== 'endret'; }
      render();
      return;
    }
    const b = el.closest('button[data-a]') as HTMLElement | null;
    if (!b || b.closest('.ov')) return;
    const a = b.getAttribute('data-a');
    const tr = b.closest('tr[data-id]');
    const o = tr ? items.filter(i => String(i.id) === tr.getAttribute('data-id'))[0] : undefined;
    if (a === 'ny') openEdit();
    else if (a === 'endre' && o) openEdit(o);
    else if (a === 'slett' && o) openDelete(o);
  });
  qInput.addEventListener('input', () => { q = qInput.value; render(); });
  ftSel.addEventListener('change', () => { fType = ftSel.value; render(); });

  const onVis = (): void => { if (document.visibilityState === 'visible' && !root.querySelector('.ov')) refresh().catch(() => undefined); };
  document.addEventListener('visibilitychange', onVis);
  const timer = window.setInterval(() => {
    if (document.visibilityState === 'visible' && !root.querySelector('.ov')) refresh().catch(() => undefined);
  }, (opts.refreshSeconds || 60) * 1000);

  render();
  refresh().catch(() => undefined);

  return {
    setDark(dark: boolean): void { root.classList.toggle('dark', dark); },
    refresh,
    dispose(): void {
      disposed = true;
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', onVis);
      host.innerHTML = '';
    }
  };
}
