/**
 * Lagring av bokføringsoppdrag i ei skjult SharePoint-liste.
 *
 * Fila er uavhengig av SPFx: ho brukar berre eit lite HTTP-grensesnitt,
 * slik at same kode kan testast mot ein simulert SharePoint utanfor Microsoft 365.
 */

export interface IOppdrag {
  id: number;
  klient: string;
  oppgavetype: string;
  periode: string;
  endret: string;   // ISO-dato
  endretAv: string;
  etag: string;
}

export type IOppdragInput = Pick<IOppdrag, 'klient' | 'oppgavetype' | 'periode'>;

export interface IHttpResponse {
  ok: boolean;
  status: number;
  json(): Promise<any>; // eslint-disable-line @typescript-eslint/no-explicit-any
  text(): Promise<string>;
}

export interface IHttp {
  get(url: string): Promise<IHttpResponse>;
  post(url: string, body: unknown, headers?: Record<string, string>): Promise<IHttpResponse>;
}

export interface IOppdragStore {
  list(): Promise<IOppdrag[]>;
  create(o: IOppdragInput): Promise<void>;
  update(o: IOppdragInput & { id: number; etag: string }): Promise<void>;
  remove(id: number): Promise<void>;
}

export class ConflictError extends Error {
  constructor() {
    super('Oppdraget er endret av en annen bruker.');
    this.name = 'ConflictError';
    Object.setPrototypeOf(this, ConflictError.prototype); // instanceof virker også når koden er kompilert til ES5
  }
}

export const LIST_TITLE: string = 'Oppdrag';
export const OPPGAVETYPER: string[] = ['Bokføring', 'MVA', 'Lønn', 'Årsoppgjør', 'Skattemelding', 'Annet'];

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&apos;');

export class SharePointListStore implements IOppdragStore {
  private _ready: Promise<void> | undefined;

  constructor(private readonly _http: IHttp, private readonly _siteUrl: string) {}

  private get _list(): string {
    return `${this._siteUrl.replace(/\/$/, '')}/_api/web/lists/GetByTitle('${LIST_TITLE}')`;
  }

  /** Lagar lista første gong appen blir brukt (krev at brukaren kan administrere lister på området). */
  public ensureList(): Promise<void> {
    if (!this._ready) {
      this._ready = this._ensure().catch(e => {
        this._ready = undefined;
        throw e;
      });
    }
    return this._ready;
  }

  private async _ensure(): Promise<void> {
    const r = await this._http.get(`${this._list}?$select=Id`);
    if (r.ok) return;
    if (r.status !== 404) throw new Error(`Kunne ikke lese listen «${LIST_TITLE}» (${r.status}).`);
    const site = this._siteUrl.replace(/\/$/, '');
    const c = await this._http.post(`${site}/_api/web/lists`, {
      Title: LIST_TITLE, BaseTemplate: 100, Hidden: true, Description: 'Data for appen Bokføringsoppdrag. Ikke rediger direkte.'
    });
    if (!c.ok) {
      throw new Error(c.status === 403
        ? 'Listen for oppdrag finnes ikke ennå, og du har ikke tilgang til å opprette den. Be en eier av området åpne appen første gang.'
        : `Kunne ikke opprette listen (${c.status}).`);
    }
    const choices = OPPGAVETYPER.map(v => `<CHOICE>${esc(v)}</CHOICE>`).join('');
    const fields = [
      `<Field Type='Choice' DisplayName='Oppgavetype' Name='Oppgavetype' Format='Dropdown' FillInChoice='TRUE'><CHOICES>${choices}</CHOICES></Field>`,
      `<Field Type='Text' DisplayName='Periode' Name='Periode' MaxLength='50' />`
    ];
    for (const xml of fields) {
      const f = await this._http.post(`${this._list}/fields/CreateFieldAsXml`, { parameters: { SchemaXml: xml, Options: 8 } });
      if (!f.ok) throw new Error(`Kunne ikke opprette felt i listen (${f.status}).`);
    }
    await this._http.post(`${this._list}/fields/GetByInternalNameOrTitle('Title')`, { Title: 'Klient' },
      { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  }

  public async list(): Promise<IOppdrag[]> {
    await this.ensureList();
    const out: IOppdrag[] = [];
    let url: string | undefined =
      `${this._list}/items?$select=Id,Title,Oppgavetype,Periode,Modified,Editor/Title&$expand=Editor&$orderby=Modified desc&$top=5000`;
    while (url) {
      const r: IHttpResponse = await this._http.get(url);
      if (!r.ok) throw new Error(`Kunne ikke hente oppdrag (${r.status}).`);
      const d = await r.json();
      (d.value || []).forEach((it: any) => out.push({ // eslint-disable-line @typescript-eslint/no-explicit-any
        id: it.Id,
        klient: it.Title || '',
        oppgavetype: it.Oppgavetype || '',
        periode: it.Periode || '',
        endret: it.Modified || '',
        endretAv: (it.Editor && it.Editor.Title) || '',
        etag: it['@odata.etag'] || it['odata.etag'] || (it.__metadata && it.__metadata.etag) || '*'
      }));
      url = d['@odata.nextLink'] || d['odata.nextLink'] || undefined;
    }
    return out;
  }

  public async create(o: IOppdragInput): Promise<void> {
    await this.ensureList();
    const r = await this._http.post(`${this._list}/items`, { Title: o.klient, Oppgavetype: o.oppgavetype, Periode: o.periode });
    if (!r.ok) throw new Error(`Kunne ikke lagre oppdraget (${r.status}).`);
  }

  public async update(o: IOppdragInput & { id: number; etag: string }): Promise<void> {
    const r = await this._http.post(`${this._list}/items(${o.id})`, { Title: o.klient, Oppgavetype: o.oppgavetype, Periode: o.periode },
      { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': o.etag || '*' });
    if (r.status === 412) throw new ConflictError();
    if (!r.ok) throw new Error(`Kunne ikke lagre endringen (${r.status}).`);
  }

  public async remove(id: number): Promise<void> {
    const r = await this._http.post(`${this._list}/items(${id})`, undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    if (!r.ok && r.status !== 404) throw new Error(`Kunne ikke slette oppdraget (${r.status}).`);
  }
}
