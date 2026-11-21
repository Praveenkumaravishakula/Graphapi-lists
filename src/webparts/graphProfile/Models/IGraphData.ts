export interface IGraphData {
  displayName: string;
  mail?: string;
  jobTitle?: string;
}

export interface ISiteData {
  id: string;
  displayName: string;
  webUrl: string;
}
// Shapes the web part shows. Graph returns more fields; we $select only these.
export interface IUserProfile {
  displayName: string;
  jobTitle?: string;
  mail?: string;
  officeLocation?: string;
  userPrincipalName: string;
}

export interface ISiteInfo {
  displayName: string;
  webUrl: string;
}

export interface ISiteList {
  id: string;
  displayName: string;
  template: string;
  lastModified: string;
  webUrl: string;
}

export type ProfileSource = 'graph' | 'pnp';