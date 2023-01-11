import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";

export interface IHttpService {
  get(query: string): Promise<any>;
  post(data: any, query: string): Promise<any>;
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

  async get(query: string): Promise<any> {
    const url = `${this._currentWebUrl}${query}`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    return response.json();
  }

  async post(data: any, query: string): Promise<any> {
    const options: ISPHttpClientOptions = {
      body: data,
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-type": "application/json;odata=verbose",
        "odata-version": "3.0",
      },
    };

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      `${this._currentWebUrl}${query}`,
      SPHttpClient.configurations.v1,
      options
    );

    return response.json();
  }
}
