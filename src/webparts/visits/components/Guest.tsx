import { ActionButton, IconButton, Separator, Stack, TextField } from "office-ui-fabric-react"
import * as React from "react"
import { FC, ReactElement } from "react"

export const GuestInfo: FC<any> = (props): ReactElement => {
    return (
        <>
            <Separator>Guest information</Separator>
            <Stack tokens={{ childrenGap: 15 }} horizontal>
                <Stack tokens={{ childrenGap: 5 }}>
                    <form onSubmit={props.guestSubmitHandler}>
                        <TextField required value={props.guest.FirstName} onChange={props.handleGuestInput} name="FirstName" label="Name" />
                        <TextField required value={props.guest.LastName} onChange={props.handleGuestInput} name="LastName" label="Last Name" />
                        <ActionButton type="submit" iconProps={{ iconName: 'AddFriend' }}>
                            Add Guest
                        </ActionButton>
                    </form>
                </Stack>
                <Stack>
                    <div>
                        <Separator>Guest list</Separator>
                        <div>
                            {props.guests?.map((guest: any, index: number) => {
                                return (<div key={index}>
                                    <span>{guest.FirstName + " " + guest.LastName}<IconButton iconProps={{ iconName: 'Cancel' }} /></span>
                                </div>)
                            })}
                        </div>
                    </div>
                </Stack>
            </Stack>
        </>
    )
}