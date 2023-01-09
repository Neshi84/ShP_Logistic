import * as React from "react";
import { FC, ReactElement, useEffect, useState } from "react";
import { IVisitsProps } from "./IVisitsProps";
import { IDropdownOption, Separator, Spinner } from "office-ui-fabric-react";
import VisitTable from "./VisitTable";
import VisitForm from "./VisitForm";
import { IOffice, IVisit, IVisitGet } from "../Types/Types";

const initialState: IVisit = {
    DateFrom: new Date(),
    DateTo: new Date(),
    HostsId: { results: [] },
    Project: "",
    Notes: "",
    OfficeId: null,
};

const Visits: FC<IVisitsProps> = (props): ReactElement => {
    // State variables
    const [visit, setVisit] = useState<IVisit>(initialState);
    const [visits, setVisits] = useState<IVisitGet[]>([]);
    const [offices, setOffices] = useState<IDropdownOption[]>([]);
    //const [guests, setGuests] = useState<IGuest[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);

    // Fetch visit list
    useEffect(() => {
        const getVisits = async (): Promise<void> => {
            const visits = await props.HttpService.get("/_api/web/lists/GetByTitle('Visits')/items?$select=ID,DateFrom,DateTo,Notes,Hosts/EMail,Hosts/FirstName,Hosts/LastName&$expand=Hosts/Id");
            console.log(visits.value);
            setVisits(visits.value);
            setLoading(false);
        }

        const getOffices = async (): Promise<void> => {
            const response = await props.HttpService.get("/_api/web/lists/GetByTitle('Offices')/items");
            const offices = response.value.map((item: IOffice) => {
                return { key: item.Id, text: item.OfficeNumber };
            });

            console.log(offices);
            setOffices(offices);
            setLoading(false);
        }

        getVisits().catch((error) => {
            console.log(error)
            setError(error)
        })

        getOffices().catch((error) => {
            console.log(error)
            setError(error)
        });

    }, []);

    return (
        <>
            {loading && <Spinner />}
            {error && <div>{error}</div>}
            {!loading && !error && (
                <>
                    <VisitForm visit={visit} setVisit={setVisit} offices={offices} context={props.context} />
                    <Separator />
                    <VisitTable
                        //guests={guests}
                        visits={visits}
                    />
                </>
            )}
        </>
    );
};

export default Visits;
