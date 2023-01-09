import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";

export interface IHttpService {
  /* getVisits(): Promise<any>; */
  get(query: string): Promise<any>;
}

export class HttpService implements IHttpService {
  public static readonly serviceKey: ServiceKey<IHttpService> =
    ServiceKey.create<IHttpService>("visit-app:IHttpService", HttpService);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);

      this._pageContext = serviceScope.consume(PageContext.serviceKey);

      //The entire PageContext object is available at this point.
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  /* getVisits(): Promise<IVisitGet> {
    const url = `${this._currentWebUrl}/_api/web/lists/GetByTitle('Visits')/items?$select=ID,DateFrom,DateTo,Notes,Hosts/EMail,Hosts/FirstName,Hosts/LastName&$expand=Hosts/Id`;

    return this._spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((items) => {
        const visits: IVisit[] = items.value.map((visit: IVisitGet) => ({
          ID: +visit.ID,
          DateFrom: visit.DateFrom,
          DateTo: visit.DateTo,
          Notes: visit.Notes,
          Hosts: visit.Hosts,
        }));
        return visits;
      })
      .catch((error) => {
        return error;
      });
  } */

  async get(query: string): Promise<any> {
    //const url = `${this._currentWebUrl}/_api/web/lists/GetByTitle('Visits')/items?$select=ID,DateFrom,DateTo,Notes,Hosts/EMail,Hosts/FirstName,Hosts/LastName&$expand=Hosts/Id`;
    const url = `${this._currentWebUrl}${query}`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    return response.json();
  }
}
