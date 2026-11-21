import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import { IUserProfile, ISiteInfo, ISiteList, ProfileSource } from '../Models/IGraphData';


const PROFILE_FIELDS = ['displayName', 'jobTitle', 'mail', 'officeLocation', 'userPrincipalName'];

export class GraphService {
  constructor(
    private _client: MSGraphClientV3, // from this.context.msGraphClientFactory.getClient('3')
    private _graph: GraphFI,          // PnPjs: graphfi().using(SPFx(this.context))
    private _siteId: string           // "{hostname},{site collection id},{web id}"
  ) { }

  // GET /me  (User.Read)
  public async getMe(source: ProfileSource): Promise<IUserProfile> {
    if (source === 'pnp') {
      const me = await this._graph.me.select(...PROFILE_FIELDS)();
      return this._toProfile(me as MicrosoftGraph.User);
    }
    const me: MicrosoftGraph.User = await this._client
      .api('/me')
      .select(PROFILE_FIELDS.join(','))
      .get();
    return this._toProfile(me);
  }

  // GET /me/photo/$value  (User.Read) -> object URL for <img>, or undefined when the user has no photo (404)
  public async getMyPhotoUrl(): Promise<string | undefined> {
    try {
      const photo: Blob = await this._client.api('/me/photo/$value').get();
      return URL.createObjectURL(photo);
    } catch (error) {
      if ((error as { statusCode?: number }).statusCode === 404) {
        return undefined;
      }
      throw error;
    }
  }

  // GET /sites/{site-id}  (Sites.Read.All)
  public async getSite(): Promise<ISiteInfo> {
    const site: MicrosoftGraph.Site = await this._client
      .api(`/sites/${this._siteId}`)
      .select('displayName,webUrl')
      .get();
    return { displayName: site.displayName || '', webUrl: site.webUrl || '' };
  }

  // GET /sites/{site-id}/lists  (Sites.Read.All) - visible lists and libraries, newest change first
  public async getLists(top: number): Promise<ISiteList[]> {
    const response: { value: MicrosoftGraph.List[] } = await this._client
      .api(`/sites/${this._siteId}/lists`)
      .select('id,displayName,list,lastModifiedDateTime,webUrl')
      .get();
    return response.value
      .filter(l => !l.list?.hidden)
      .sort((a, b) => (b.lastModifiedDateTime || '').localeCompare(a.lastModifiedDateTime || ''))
      .slice(0, top)
      .map(l => ({
        id: l.id || '',
        displayName: l.displayName || '',
        template: l.list?.template || '',
        lastModified: l.lastModifiedDateTime || '',
        webUrl: l.webUrl || ''
      }));
  }

  private _toProfile(me: MicrosoftGraph.User): IUserProfile {
    return {
      displayName: me.displayName || '',
      jobTitle: me.jobTitle || undefined,
      mail: me.mail || undefined,
      officeLocation: me.officeLocation || undefined,
      userPrincipalName: me.userPrincipalName || ''
    };
  }
}

// Turns a Graph error (MSGraphClientV3 GraphError or PnPjs HttpRequestError) into a readable line.
export function graphErrorText(error: unknown): string {
  const e = error as { statusCode?: number; status?: number; code?: string; message?: string };
  const status = e.statusCode || e.status;
  return [status ? `HTTP ${status}` : '', e.code || '', e.message || String(error)].filter(Boolean).join(' - ');
}

