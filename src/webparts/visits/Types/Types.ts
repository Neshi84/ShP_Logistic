import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";
import { IHttpService } from "../services/HttpService";

export interface IVisitsProps {
  context: WebPartContext;
  HttpService?: IHttpService;
}

export interface IGuest {
  ID?: string;
  VisitId?: number;
  FirstName: string;
  LastName: string;
}

export interface IOffice {
  Id: number;
  OfficeNumber: string;
}

export interface IHostsId {
  results: number[];
}

export interface IVisit {
  DateFrom: Date;
  DateTo: Date;
  Notes: string;
  HostsId: IHostsId;
  Project: string;
  OfficeId: number;
}

export interface IHost {
  EMail: string;
  FirstName: string;
  LastName: string;
}

export interface IVisitGet {
  ID?: number;
  DateFrom: Date;
  DateTo: Date;
  Notes: string;
  Hosts: IHost[];
  Project: string;
  OfficeId: number;
}

export interface IUser {
  id: number;
  text: string;
  secondaryText: string;
}

export interface IVisitFormProps {
  visit: IVisit;
  setVisit: React.Dispatch<React.SetStateAction<IVisit>>;
  HttpService: IHttpService;
  context: WebPartContext;
  offices: IDropdownOption[];
}

export interface IVisitTableProps {
  visits: IVisitGet[];
  getVisitGuests: (visitId: number) => IGuest[];
}
