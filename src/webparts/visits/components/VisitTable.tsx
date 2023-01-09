import { Persona, PersonaSize, Stack } from "office-ui-fabric-react";
import * as React from "react";
import { FC } from "react";
import { formatDate } from "../Helpers/Helpers";
import { IVisitsProps } from "../Types/Types";
import styles from "./Visits.module.scss";

const VisitTable: FC<IVisitsProps> = (props): React.ReactElement => {
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
                    {props.visits.map((visit) => {
                        return (
                            <tr key={visit.ID}>
                                <td>{visit.ID}</td>
                                <td>{formatDate(visit.DateFrom)}</td>
                                <td>{formatDate(visit.DateTo)}</td>
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
