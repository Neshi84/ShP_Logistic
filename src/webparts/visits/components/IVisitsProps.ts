import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IHttpService } from "../services/HttpService";

export interface IVisitsProps {
  context: WebPartContext;
  HttpService?: IHttpService;
}
