import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, type ISPHttpClientOptions } from '@microsoft/sp-http';

import * as strings from 'OppdragWebPartStrings';
import { type IHttp, SharePointListStore } from './core/store';
import { type IAppHandle, mountOppdragApp } from './core/ui';

export interface IOppdragWebPartProps {
  siteUrl: string;
}

/** Felles innstilling for hele tenanten: hvilket område listen ligger på. */
const STORAGE_KEY: string = 'OppdragDataSite';

export default class OppdragWebPart extends BaseClientSideWebPart<IOppdragWebPartProps> {
  private _app: IAppHandle | undefined;
  private _dark: boolean = false;
  private _siteUrl: string = '';

  private _http(): IHttp {
    const client = this.context.spHttpClient;
    const base = {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
    return {
      get: url => client.get(url, SPHttpClient.configurations.v1, { headers: base }),
      post: (url, body, headers) => {
        const o: ISPHttpClientOptions = { headers: { ...base, ...(headers || {}) } };
        if (body !== undefined) o.body = JSON.stringify(body);
        return client.post(url, SPHttpClient.configurations.v1, o);
      }
    };
  }

  /** Egenskap på web-delen → tenant-innstilling → området appen kjører på. */
  private async _resolveSite(): Promise<string> {
    if (this.properties.siteUrl) return this.properties.siteUrl.replace(/\/$/, '');
    const web = this.context.pageContext.web.absoluteUrl;
    try {
      const r = await this._http().get(`${web}/_api/web/GetStorageEntity('${STORAGE_KEY}')`);
      if (r.ok) {
        const d = await r.json();
        if (d && d.Value) return String(d.Value).replace(/\/$/, '');
      }
    } catch {
      // ingen tenant-innstilling
    }
    return web;
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    const teams = this.context.sdks.microsoftTeams;
    if (teams) {
      try {
        const ctx = await teams.teamsJs.app.getContext();
        this._dark = ctx.app.theme === 'dark' || ctx.app.theme === 'contrast';
        teams.teamsJs.app.registerOnThemeChangeHandler(t => {
          this._dark = t === 'dark' || t === 'contrast';
          if (this._app) this._app.setDark(this._dark);
        });
      } catch {
        // ikke i Teams
      }
    }
    this._siteUrl = await this._resolveSite();
  }

  public render(): void {
    if (this._app) return;
    this.domElement.style.height = '100%';
    const store = new SharePointListStore(this._http(), this._siteUrl);
    this._app = mountOppdragApp(this.domElement, store, { dark: this._dark });
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    if (!theme || this.context.sdks.microsoftTeams) return;
    this._dark = !!theme.isInverted;
    if (this._app) this._app.setDark(this._dark);
  }

  protected onPropertyPaneConfigurationComplete(): void {
    this._resolveSite().then(url => {
      this._siteUrl = url;
      this.onDispose();
      this.render();
    }).catch(() => undefined);
  }

  protected onDispose(): void {
    if (this._app) this._app.dispose();
    this._app = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [PropertyPaneTextField('siteUrl', { label: strings.SiteUrlFieldLabel, description: strings.SiteUrlFieldHelp })]
        }]
      }]
    };
  }
}
