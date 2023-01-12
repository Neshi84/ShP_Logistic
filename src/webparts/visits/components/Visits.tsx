import * as React from "react";
import { FC, ReactElement, useEffect, useState } from "react";
import { IDropdownOption, Separator, Spinner } from "office-ui-fabric-react";
import VisitTable from "./VisitTable";
import VisitForm from "./VisitForm";
import { IGuest, IOffice, IVisit, IVisitGet, IVisitsProps } from "../Types/Types";

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
    const [guests, setGuests] = useState<IGuest[]>([]);
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

        const getGuests = async (): Promise<void> => {
            const response = await props.HttpService.get("/_api/web/lists/GetByTitle('Guests')/items");
            setGuests(response.value);
            console.log(response.value);
        }

        const getOffices = async (): Promise<void> => {
            const response = await props.HttpService.get("/_api/web/lists/GetByTitle('Offices')/items");
            const offices = response.value.map((item: IOffice) => {
                return { key: item.Id, text: item.OfficeNumber };
            });
            setOffices(offices);
            setLoading(false);
        }

        getVisits().catch((error) => {
            console.log(error)
            setError(error)
        })

        getGuests().catch((error) => {
            console.log(error)
            setError(error)
        })

        getOffices().catch((error) => {
            console.log(error)
            setError(error)
        });

    }, []);

    const getVisitGuests = (visitId: number): IGuest[] => {
        const filteredGuests = guests.filter(guest => guest.VisitId === visitId);
        console.log(filteredGuests);
        return filteredGuests
    }

    return (
        <>
            {loading && <Spinner />}
            {error && <div>{error}</div>}
            {!loading && !error && (
                <>
                    <VisitForm
                        visit={visit}
                        setVisit={setVisit}
                        offices={offices}
                        HttpService={props.HttpService}
                        context={props.context}
                    />
                    <Separator />
                    <VisitTable
                        //guests={guests}
                        visits={visits}
                        getVisitGuests={getVisitGuests}
                    />
                </>
            )}
        </>
    );
};

export default Visits;

