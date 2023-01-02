import * as React from "react";
import { FC, ReactElement, useEffect, useState } from "react";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
  DateTimePicker,
  DateConvention,
  TimeConvention,
  TimeDisplayControlType,
} from "@pnp/spfx-controls-react/lib/DateTimePicker";
import { IVisitsProps } from "./IVisitsProps";
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IStackTokens,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import { IIconProps } from '@fluentui/react';
import { ActionButton } from '@fluentui/react/lib/Button';
import VisitTable from "./VisitTablet";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 15,
};

const addFriendIcon: IIconProps = { iconName: 'AddFriend' };

// Interfaces
interface IVisit {
  DateFrom: Date;
  DateTo: Date;
  Notes: string;
  HostsId: { results: number[] };
  Project: string;
  OfficeId: number;
}

const initialState: IVisit = {
  DateFrom: new Date(),
  DateTo: new Date(),
  HostsId: { results: [] },
  Project: "",
  Notes: "",
  OfficeId: null
}

interface IOffice {
  Id: number;
  OfficeNumber: string;
}

/* interface IGuest {
  VisitID: number;
  FirstName: string;
  LastName: string;
} */

const Visits: FC<IVisitsProps> = (props): ReactElement => {

  const [visit, setVisit] = useState<IVisit>(initialState);
  const [offices, setOffices] = useState<IDropdownOption[]>([])
  //const [guests, setGuests] = useState<IGuest[]>([]);

  useEffect(() => {
    const getListData = async (): Promise<void> => {
      const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Offices')/items`;
      const response: SPHttpClientResponse =
        await props.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1
        );
      const listData: any = await response.json();

      const items = listData.value.map((item: IOffice) => {
        return { key: item.Id, text: item.OfficeNumber };
      });

      setOffices(items);
    };

    getListData().catch(console.error);
  }, []);

  const getPeoplePickerItems = (items: any[]): void => {
    setVisit({ ...visit, HostsId: { results: [...items.map((x) => x.id)] } });
  };

  const handleChangeDateFrom = (date: Date): void => {
    setVisit({ ...visit, DateFrom: date });
  };

  const handleChangeDateTo = (date: Date): void => {
    setVisit({ ...visit, DateTo: date });
  };

  const handleInput = (
    event: React.FormEvent<
      HTMLInputElement | HTMLTextAreaElement | HTMLDivElement
    >
  ): void => {
    console.log(event.target);
    const target = event.target as HTMLInputElement | HTMLTextAreaElement;
    console.log(target.value);
    setVisit({ ...visit, [target.name]: target.value })
  };

  const handleOnChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    console.log(item);
    setVisit({ ...visit, OfficeId: +item.key })
  };

  const handleSubmit = (): void => {

    const data = { ...visit, __metadata: { type: "SP.Data.VisitsListItem" } };
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(data),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-type": "application/json;odata=verbose",
        "odata-version": "3.0",
      },
    };

    props.context.spHttpClient
      .post(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Visits')/items`,
        SPHttpClient.configurations.v1,
        options
      )
      .then((response: SPHttpClientResponse) => {
        response
          .json()
          .then((data) => {
            const guests = [
              {
                VisitId: data.d.ID,
                Name: "Nebojsa",
                LastName: "Milovac",
                __metadata: { type: "SP.Data.GuestsListItem" },
              },
              {
                VisitId: data.d.ID,
                Name: "Test1",
                LastName: "Test1Last",
                __metadata: { type: "SP.Data.GuestsListItem" },
              },
              {
                VisitId: data.d.ID,
                Name: "Nebojsa2",
                LastName: "Milovac2",
                __metadata: { type: "SP.Data.GuestsListItem" },
              },
              {
                VisitId: data.d.ID,
                Name: "Nebojsa3",
                LastName: "Milovac3",
                __metadata: { type: "SP.Data.GuestsListItem" },
              }
            ];

            guests.forEach(guest => {

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
                .then((data) => {
                  console.log(data);
                })
                .catch((error) => {
                  console.log(error);
                });

            })
          })
          .catch((error) => {
            console.log(error);
          });
      })
      .catch((error) => {
        console.log(error);
      });
  };

  return (
    <>
      <section>
        <Stack tokens={{ childrenGap: 15 }}>
          <PeoplePicker
            context={props.context}
            titleText="Hosts"
            showtooltip={true}
            required={true}
            onChange={getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            personSelectionLimit={3}
            resolveDelay={1000}
            ensureUser={true}
            defaultSelectedUsers={[]}
          />
          <Stack horizontal tokens={itemAlignmentsStackTokens}>
            <DateTimePicker
              label="Visit from: "
              timeDisplayControlType={TimeDisplayControlType.Dropdown}
              showLabels={false}
              dateConvention={DateConvention.DateTime}
              timeConvention={TimeConvention.Hours24}
              minDate={new Date()}
              value={visit.DateFrom}
              onChange={handleChangeDateFrom}
            />
            <DateTimePicker
              label="Visit to: "
              timeDisplayControlType={TimeDisplayControlType.Dropdown}
              showLabels={false}
              dateConvention={DateConvention.DateTime}
              timeConvention={TimeConvention.Hours24}
              minDate={visit.DateFrom}
              value={visit.DateTo}
              onChange={handleChangeDateTo}
            />
          </Stack>
          <Stack tokens={{ childrenGap: 5 }}>
            <TextField name="Project" onChange={handleInput} label="Project" />
            <Dropdown
              onChange={handleOnChange}
              label="Offices"
              options={offices}
              placeholder="Select an office"
            />
            <TextField
              name="Notes"
              onChange={handleInput}
              multiline
              rows={3}
              label="Notes"
            />
          </Stack>
          <Stack tokens={{ childrenGap: 15 }} horizontal>
            <Stack tokens={{ childrenGap: 5 }}>
              <TextField label="Name" />
              <TextField label="Last Name" />
              <ActionButton iconProps={addFriendIcon}>
                Add Guest
              </ActionButton>
            </Stack>
            <Stack>
              <div>Guest List</div>
            </Stack>
          </Stack>
          <DefaultButton text="Save" onClick={handleSubmit} />
        </Stack>
      </section>
      <section>
        <Separator />
        <VisitTable context={props.context} />
      </section>
    </>
  );
};

export default Visits;
