import { DateConvention, DateTimePicker, TimeConvention, TimeDisplayControlType } from "@pnp/spfx-controls-react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ActionButton, DefaultButton, Dropdown, IconButton, IDropdownOption, IPersonaProps, MessageBar, MessageBarType, Separator, Stack, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { FC, ReactElement, useState } from "react";
import { IGuest, IVisit, IVisitFormProps } from "../Types/Types";

const guestInitialState: IGuest = {
    FirstName: "",
    LastName: ""
}

const VisitForm: FC<IVisitFormProps> = (props): ReactElement => {
    // State variables
    const [status, setStatus] = useState<"success" | "error" | null>(
        null
    );
    const [visit, setVisit] = useState<IVisit>(props.visit);
    const [guests, setGuests] = useState<IGuest[]>([]);
    const [guest, setGuest] = useState<IGuest>(guestInitialState);


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
        setGuest(guestInitialState);
    }

    // Event handlers
    const getPeoplePickerItems = (items: IPersonaProps[]): void => {
        setVisit({ ...visit, HostsId: { results: [...items.map((x) => +x.id)] } });
    };

    const handleDateChange = (field: string, date: Date): void => {
        setVisit({ ...visit, [field]: date });
    }

    const handleInput = (
        event: React.FormEvent<
            HTMLInputElement | HTMLTextAreaElement
        >
    ): void => {
        const target = event.target as HTMLInputElement | HTMLTextAreaElement;
        setVisit({ ...visit, [target.name]: target.value });
    };

    const handleOnChange = (
        _event: React.FormEvent<HTMLDivElement>,
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
        setGuests([...guests, guest]);
        setGuest(guestInitialState);
    }

    const handleSubmit = async (): Promise<void> => {
        try {
            // Validate input
            if (!visit.DateFrom || !visit.DateTo || !visit.HostsId.results.length) {
                setStatus("error");
                return;
            }

            const data = JSON.stringify({ ...visit, __metadata: { type: "SP.Data.VisitsListItem" } });
            const visitData = await props.HttpService.post(data, "/_api/web/lists/GetByTitle('Visits')/items");

            const guestData = guests.map(guest => {
                return ({ ...guest, VisitId: visitData.d.ID, __metadata: { type: "SP.Data.GuestsListItem" } })
            })

            for (const guest of guestData) {
                const data = JSON.stringify(guest);
                const response = await props.HttpService.post(data, "/_api/web/lists/GetByTitle('Guests')/items");
                console.log(response.d);
            }

            await props.getVisits();
            resetForm();
            setStatus("success");

        } catch (error) {
            console.log(error)
            setStatus("error");
        }
    };

    return (
        <>
            <Separator>Visit information</Separator>
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
                <Stack horizontal tokens={{ childrenGap: 15 }}>
                    <DateTimePicker
                        label="From:"
                        value={visit.DateFrom}
                        key={"DateFrom"}
                        onChange={date => handleDateChange('DateFrom', date)}
                        dateConvention={DateConvention.DateTime}
                        timeConvention={TimeConvention.Hours12}
                        isMonthPickerVisible={true}
                        showLabels={false}
                        timeDisplayControlType={TimeDisplayControlType.Dropdown}
                    />
                    <DateTimePicker
                        label="To:"
                        value={visit.DateTo}
                        key={"DateTo"}
                        onChange={date => handleDateChange('DateTo', date)}
                        dateConvention={DateConvention.DateTime}
                        timeConvention={TimeConvention.Hours12}
                        isMonthPickerVisible={true}
                        showLabels={false}
                        timeDisplayControlType={TimeDisplayControlType.Dropdown}
                    />
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 15 }}>
                    <Stack.Item grow={5}>
                        <TextField
                            label="Project"
                            placeholder="Project"
                            required={true}
                            name="Project"
                            value={visit.Project}
                            onChange={handleInput}
                        />
                    </Stack.Item>
                    <Stack.Item grow={1}>
                        <Dropdown
                            label="Office"
                            placeholder="Office"
                            options={props.offices}
                            onChange={handleOnChange}
                            required={true}
                            selectedKey={visit.OfficeId}
                        />
                    </Stack.Item>
                </Stack>
                <TextField
                    label="Reason for visit"
                    placeholder="Reason for visit"
                    required={true}
                    multiline
                    rows={3}
                    name="Notes"
                    value={visit.Notes}
                    onChange={handleInput}
                />
                <Separator>Guest information</Separator>
                <Stack tokens={{ childrenGap: 15 }} horizontal>
                    <Stack horizontal tokens={{ childrenGap: 5 }}>
                        <form onSubmit={guestSubmitHandler}>
                            <Stack horizontal tokens={{ childrenGap: 5 }}>
                                <TextField required value={guest.FirstName} onChange={handleGuestInput} name="FirstName" label="Name" />
                                <TextField required value={guest.LastName} onChange={handleGuestInput} name="LastName" label="Last Name" />
                            </Stack>
                            <ActionButton type="submit" iconProps={{ iconName: 'AddFriend' }}>
                                Add Guest
                            </ActionButton>
                        </form>
                    </Stack>
                    <Stack>
                        <div>
                            {guests.length > 0 &&
                                <Separator alignContent="start">Guest list</Separator>}
                            <div>
                                {guests?.map((guest: any, index: number) => {
                                    return (<div key={index}>
                                        <span>{guest.FirstName + " " + guest.LastName}<IconButton iconProps={{ iconName: 'Cancel' }} /></span>
                                    </div>)
                                })}
                            </div>
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