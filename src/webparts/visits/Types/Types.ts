export interface IGuest {
  VisitID?: number;
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

export interface IVisitsProps {
  visits: IVisitGet[];
}
