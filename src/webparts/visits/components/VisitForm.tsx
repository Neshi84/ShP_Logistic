import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { DateConvention, DateTimePicker, TimeConvention, TimeDisplayControlType } from "@pnp/spfx-controls-react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ActionButton, DefaultButton, Dropdown, IconButton, IDropdownOption, MessageBar, MessageBarType, Stack, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { FC, ReactElement, useState } from "react";
import { IGuest, IVisit } from "../Types/Types";

const VisitForm: FC<any> = (props): ReactElement => {
    // State variables
    const [status, setStatus] = useState<"success" | "error" | null>(
        null
    );
    const [visit, setVisit] = useState<IVisit>(props.visit);
    const [guests, setGuests] = useState<IGuest[]>([]);
    const [guest, setGuest] = useState<IGuest>(null);


    const resetForm = (): void => {
        setVisit({
            DateFrom: new Date(),
            DateTo: new Date(),
            HostsId: { results: [] },
            Project: "",
            Notes: "",
            OfficeId: null,
        });
        setGuests([]);
        setGuest(null);

    }

    // Event handlers
    const getPeoplePickerItems = (items: any[]): void => {
        console.log(items)
        setVisit({ ...visit, HostsId: { results: [...items.map((x) => x.id)] } });
    };

    const handleDateChange = (field: string, date: Date): void => {
        setVisit({ ...visit, [field]: date });
    }

    const handleInput = (
        event: React.FormEvent<
            HTMLInputElement | HTMLTextAreaElement | HTMLDivElement
        >
    ): void => {
        const target = event.target as HTMLInputElement | HTMLTextAreaElement;
        setVisit({ ...visit, [target.name]: target.value });
    };

    const handleOnChange = (
        event: React.FormEvent<HTMLDivElement>,
        item: IDropdownOption
    ): void => {

        setVisit({ ...visit, OfficeId: +item.key });
    };

    const handleGuestInput = (
        event: React.FormEvent<
            HTMLInputElement
        >
    ): void => {
        const target = event.target as HTMLInputElement | HTMLTextAreaElement;
        setGuest({ ...guest, [target.name]: target.value })
    };

    const guestSubmitHandler = (event: React.FormEvent<HTMLFormElement>): void => {
        event.preventDefault();
        console.log(guest)
        console.log([...guests, guest])
        setGuests([...guests, guest]);
    }

    const handleSubmit = async (): Promise<void> => {
        try {
            // Validate input
            if (!visit.DateFrom || !visit.DateTo || !visit.HostsId.results.length) {
                setStatus("error");
                return;
            }

            const options: ISPHttpClientOptions = {
                body: JSON.stringify({ ...visit, __metadata: { type: "SP.Data.VisitsListItem" } }),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-type": "application/json;odata=verbose",
                    "odata-version": "3.0"
                }
            };

            const response: SPHttpClientResponse = await props.context.spHttpClient.post(
                `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Visits')/items`,
                SPHttpClient.configurations.v1,
                options
            );

            if (response.ok) {
                const visit = await response.json();
                const guestData = guests.map(guest => {
                    return ({ ...guest, VisitId: visit.d.ID, __metadata: { type: "SP.Data.GuestsListItem" } })
                })

                guestData.forEach(guest => {
                    const options: ISPHttpClientOptions = {
                        body: JSON.stringify(guest),
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "Content-type": "application/json;odata=verbose",
                            "odata-version": "3.0",
                        },
                    };

                    props.context.spHttpClient
                        .post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Guests')/items`,
                            SPHttpClient.configurations.v1,
                            options
                        )
                        .then((data: SPHttpClientResponse) => {
                            console.log(data);
                        })
                        .catch((error: SPHttpClientResponse) => {
                            console.log(error);
                        });
                })

                resetForm();
                setStatus("success");

            } else {
                console.log(response.json().then(data => console.log(data.error.message)))
                setStatus("error");
            }
        } catch (error) {
            setStatus("error");
        }
    };

    return (
        <>
            {status === "error" && (
                <MessageBar messageBarType={MessageBarType.error}>Please fill out all required fields before submitting.</MessageBar>
            )}
            <Stack tokens={{ childrenGap: 15 }}>
                <PeoplePicker
                    context={props.context}
                    titleText="Hosts"
                    personSelectionLimit={3}
                    showtooltip={true}
                    required={true}
                    disabled={false}
                    ensureUser={true}
                    onChange={getPeoplePickerItems}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers={[]}
                />
                <DateTimePicker
                    label="From"
                    value={visit.DateFrom}
                    key={"DateFrom"}
                    onChange={date => handleDateChange('DateFrom', date)}
                    dateConvention={DateConvention.DateTime}
                    timeConvention={TimeConvention.Hours12}
                    showLabels={true}
                    isMonthPickerVisible={true}
                    timeDisplayControlType={TimeDisplayControlType.Dropdown}
                />
                <DateTimePicker
                    label="To"
                    value={visit.DateTo}
                    key={"DateTo"}
                    onChange={date => handleDateChange('DateTo', date)}
                    dateConvention={DateConvention.DateTime}
                    timeConvention={TimeConvention.Hours12}
                    showLabels={true}
                    isMonthPickerVisible={true}
                    timeDisplayControlType={TimeDisplayControlType.Dropdown}
                />
                <TextField
                    label="Project"
                    required={true}
                    name="Project"
                    value={visit.Project}
                    onChange={handleInput}
                />
                <TextField
                    label="Notes"
                    required={true}
                    name="Notes"
                    value={visit.Notes}
                    onChange={handleInput}
                />
                <Dropdown
                    label="Office"
                    options={props.offices}
                    onChange={handleOnChange}
                    required={true}
                    selectedKey={visit.OfficeId}
                />

                <Stack tokens={{ childrenGap: 15 }} horizontal>
                    <Stack tokens={{ childrenGap: 5 }}>
                        <form onSubmit={guestSubmitHandler}>
                            <TextField required onChange={handleGuestInput} name="FirstName" label="Name" />
                            <TextField required onChange={handleGuestInput} name="LastName" label="Last Name" />
                            <ActionButton type="submit" iconProps={{ iconName: 'AddFriend' }}>
                                Add Guest
                            </ActionButton>
                        </form>
                    </Stack>
                    <Stack>
                        <div>
                            <h3>Guest List</h3>
                            <ul>
                                {guests?.map(guest => {
                                    return (<li key={guest.FirstName + guest.LastName}>
                                        <span>{guest.FirstName + " " + guest.LastName}<IconButton iconProps={{ iconName: 'Cancel' }} /></span>
                                    </li>)
                                })}
                            </ul>
                        </div>
                    </Stack>
                </Stack>

                {status === "success" &&
                    <MessageBar
                        messageBarType={MessageBarType.success}
                        isMultiline={false}
                    >
                        Visit is saved succesfully!
                    </MessageBar>
                }
                <DefaultButton
                    text="Submit"
                    onClick={handleSubmit}
                />
            </Stack>

        </>

    );
};
export default VisitForm;