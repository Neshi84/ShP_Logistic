import * as React from "react";
import { FC, useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IVisitsProps } from "./IVisitsProps";
import { Persona, PersonaSize, Stack } from "office-ui-fabric-react";
import styles from "./Visits.module.scss";


interface IHost {
    EMail: string;
    FirstName: string;
    LastName: string;
}

interface IVisit {
    ID: number;
    DateFrom: string;
    DateTo: string;
    Notes: string;
    Hosts: IHost[];
}

const VisitTable: FC<IVisitsProps> = (props): React.ReactElement => {
    const [visits, setVisitis] = useState<IVisit[]>([]);

    useEffect(() => {
        getVisits();
    }, []);

    const formatDate = (date: Date): string => {
        const minutes = date.getMinutes();
        const formatedMinutes = minutes < 10 ? "0" + minutes : minutes;
        const strTime = date.getHours() + ":" + formatedMinutes;
        return `${date.getDate()}.${date.getMonth() + 1
            }.${date.getFullYear()}  ${strTime}`;
    };

    function getVisits(): void {
        const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Visits')/items?$select=ID,DateFrom,DateTo,Notes,Hosts/EMail,Hosts/FirstName,Hosts/LastName&$expand=Hosts/Id`;

        props.context.spHttpClient
            .get(url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => response.json())
            .then((items) => {
                const visits: IVisit[] = items.value.map((visit: IVisit) => ({
                    ID: +visit.ID,
                    DateFrom: visit.DateFrom,
                    DateTo: visit.DateTo,
                    Notes: visit.Notes,
                    Hosts: visit.Hosts,
                }));
                console.log(visits);
                setVisitis(visits);
            })
            .catch((error) => {
                console.log(error);
            });
    }

    return (
        <>
            <div>Visits</div>
            <div>
                <table className={styles.visitTable}>
                    <tr>
                        <th>Id</th>
                        <th>Visit From</th>
                        <th>Visit To</th>
                        <th>Notes</th>
                        <th>Hosts</th>
                    </tr>
                    {visits.map((visit) => {
                        return (
                            <tr key={visit.ID}>
                                <td>{visit.ID}</td>
                                <td>{formatDate(new Date(visit.DateFrom))}</td>
                                <td>{formatDate(new Date(visit.DateTo))}</td>
                                <td>{visit.Notes}</td>
                                <td>
                                    <Stack horizontal>
                                        {visit.Hosts.map((x) => {
                                            return (
                                                <Persona
                                                    key={x.EMail}
                                                    size={PersonaSize.size24}
                                                    text={`${x.FirstName} ${x.LastName}`}
                                                    secondaryText={x.EMail}
                                                />
                                            );
                                        })}
                                    </Stack>
                                </td>
                            </tr>
                        );
                    })}
                </table>
            </div>
        </>
    );
};

export default VisitTable;
