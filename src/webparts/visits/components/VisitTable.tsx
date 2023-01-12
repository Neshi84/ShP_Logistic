import { Persona, PersonaSize, Separator, Stack } from "office-ui-fabric-react";
import * as React from "react";
import { FC } from "react";
import { formatDate } from "../Helpers/Helpers";
import { IGuest, IHost, IVisitGet, IVisitTableProps } from "../Types/Types";
import styles from "./Visits.module.scss";

const VisitTable: FC<IVisitTableProps> = (props): React.ReactElement => {
    return (
        <>
            <Separator>Visit list</Separator>
            <div>
                <table className={styles.visitTable}>
                    <tr>
                        <th>Id</th>
                        <th>Visit From</th>
                        <th>Visit To</th>
                        <th>Notes</th>
                        <th>Hosts</th>
                        <th>Guests</th>
                    </tr>
                    {props.visits.map((visit: IVisitGet) => {
                        return (
                            <tr key={visit.ID}>
                                <td>{visit.ID}</td>
                                <td>{formatDate(visit.DateFrom)}</td>
                                <td>{formatDate(visit.DateTo)}</td>
                                <td>{visit.Notes}</td>
                                <td>
                                    <Stack tokens={{ childrenGap: 5 }}>
                                        {visit.Hosts.map((x: IHost) => {
                                            return (
                                                <Persona
                                                    key={x.EMail}
                                                    size={PersonaSize.size24}
                                                    text={`${x.FirstName} ${x.LastName}`}
                                                />
                                            );
                                        })}
                                    </Stack>
                                </td>
                                <td>
                                    <Stack tokens={{ childrenGap: 5 }}>
                                        {props.getVisitGuests(visit.ID).map((guest: IGuest) => {
                                            return (
                                                <Persona
                                                    key={guest.ID}
                                                    size={PersonaSize.size24}
                                                    text={`${guest.FirstName} ${guest.LastName}`}
                                                />
                                            )
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
