import * as React from "react";
import styles from "./UserIsOwner.module.scss";
import { IUserIsOwnerProps } from "./IUserIsOwnerProps";
import { Toggle } from "@fluentui/react/lib/Toggle";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";
import { PrimaryButton } from "@fluentui/react";
import { Persona } from "office-ui-fabric-react/lib/Persona";
import { Stack } from "office-ui-fabric-react/lib/components/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { List } from "office-ui-fabric-react/lib/List";
import { Label } from "office-ui-fabric-react/lib/Label";
import { FocusZone, TextField } from "@fluentui/react";

interface IUser {
  id: string;
  name: string;
  mail: string;
}

interface IGroupResponse {
  "@odata.context": string;
  value: any[];
}

interface IUserIsOwnerState {
  listItems: any[];
  filteredItems: any[];
  showPrivate: boolean;
  user: IUser;
}

export default class UserIsOwner extends React.Component<
  IUserIsOwnerProps,
  IUserIsOwnerState
> {
  constructor(props: IUserIsOwnerProps) {
    super(props);
    this.state = {
      listItems: [],
      filteredItems: [],
      showPrivate: false,
      user: {
        id: "",
        mail: "",
        name: "",
      },
    };

    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getUserOwnerGroups = this._getUserOwnerGroups.bind(this);
    this._setUserOwnerGroups = this._setUserOwnerGroups.bind(this);
    this._myGroups = this._myGroups.bind(this);
    this._onChange = this._onChange.bind(this);
    
  }

  private _getPeoplePickerItems(items: any[]): void {
    console.log("Items", items);
    const selectedUser = items[0];
    console.log("Selected user", selectedUser);
    this.setState(
      {
        user: {
          id: selectedUser.id,
          mail: selectedUser.secondaryText,
          name: selectedUser.text,
        },
      },
      () => this._getUserOwnerGroups()
    );
  }

  private _getUserOwnerGroups(): void {
    const groups = this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/users/${this.state.user.mail}/ownedObjects`)
          .get((error, response) => {
            if (error) {
              console.log(error);
            }
            console.log(response);
            //filter out groups that have @odata.type === "#microsoft.graph.group"
            response.value = response.value.filter((group: any) => {
              return group["@odata.type"] === "#microsoft.graph.group";
            });
            this.setState({
              listItems: response.value,
              filteredItems: response.value,
            });
          })
          .catch((error) => {
            console.log(error);
          });
      })
      .catch((error) => {
        console.log(error);
      });
    console.log(groups);

    this._setUserOwnerGroups();
  }

  private _setUserOwnerGroups(): void {
    let listItems = this.state.listItems;
    console.log("listItems", listItems);

    //filter out private groups if showPrivate is false
    if (!this.state.showPrivate) {
      listItems = listItems.filter((group: any) => {
        return group.visibility === "Public";
      });
    }

    //sort the groups by displayName
    listItems.sort((a: any, b: any) => {
      if (a.displayName.toLowerCase() < b.displayName.toLowerCase()) {
        return -1;
      }
      if (a.displayName.toLowerCase() > b.displayName.toLowerCase()) {
        return 1;
      }
      return 0;
    });

    this.setState({
      listItems: listItems,
      filteredItems: listItems,
    });
  }

  private _myGroups(): void {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/me/ownedObjects`)
          .get((error, response) => {
            if (error) {
              console.log(error);
            }
            console.log(response);
          })
          .catch((error) => {
            console.log(error);
          });
      })
      .catch((error) => {
        console.log(error);
      });
  }

  private _onChange(
    ev: React.MouseEvent<HTMLElement>,
    isChecked?: boolean
  ): void {
    this.setState({
      showPrivate: isChecked,
    });
    this._getUserOwnerGroups();
  }

  public render(): React.ReactElement<IUserIsOwnerProps> {
    const { user, listItems, filteredItems } = this.state;
    const { hasTeamsContext, context } = this.props;
    const onRenderCell = (item: any): JSX.Element => {
      return (
        <li data-is-focusable className={styles.listItem} key={item.id}>
          <LivePersona
            upn={item.mail}
            template={
              <>
                <Persona
                  text={item.displayName}
                  secondaryText={item.mail}
                  coinSize={32}
                />
              </>
            }
            serviceScope={this.context.serviceScope}
          />
          {item.description && (
            <Label style={{ marginTop: "5px" }}>{item.description}</Label>
          )}
        </li>
      );
    };
    const stackTokens = { childrenGap: 20 };

    const resultCountText =
      filteredItems.length === listItems.length
        ? ""
        : ` (${filteredItems.length} of ${listItems.length} shown)`;

    const onFilterChanged = (_: any, text: string): void => {
      const newFilteredItems = listItems.filter(
        (item) =>
          item.displayName.toLowerCase().indexOf(text.toLowerCase()) >= 0
      );
      this.setState({ filteredItems: newFilteredItems });
    };

    return (
      <section
        className={`${styles.userIsOwner} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <Stack tokens={stackTokens}>
          <Text variant="xLarge">Group ownerships</Text>
          <Stack horizontal tokens={stackTokens}>
            <Stack.Item grow={1}>
              <PeoplePicker
                context={context as any}
                titleText="Select a user"
                personSelectionLimit={1}
                showtooltip={true}
                onChange={this._getPeoplePickerItems.bind(this)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
            </Stack.Item>
            <Stack.Item align="end">
              {user.mail.length > 0 && (
                <LivePersona
                  upn={user.mail}
                  template={
                    <>
                      <Persona
                        text={user.name}
                        secondaryText={user.mail}
                        coinSize={24}
                      />
                    </>
                  }
                  serviceScope={this.context.serviceScope}
                />
              )}
            </Stack.Item>
          </Stack>
          <Stack tokens={stackTokens}>
            {listItems.length !== 0 && (
              <Stack.Item align="stretch" grow={1}>
                <TextField
                  label={"Filter by name" + resultCountText}
                  onChange={onFilterChanged}
                />
                <Toggle
                  label="Show private groups"
                  onText="On"
                  offText="Off"
                  onChange={this._onChange}
                />
              </Stack.Item>
            )}
            {listItems.length === 0 && user.mail.length > 0 && (
              <Stack.Item>
                <Label>There are no groups owned by this user</Label>
              </Stack.Item>
            )}
            <Stack.Item align="stretch">
              <List
                className={styles.list}
                items={filteredItems}
                onRenderCell={onRenderCell}
              />
            </Stack.Item>
          </Stack>
        </Stack>
      </section>
    );
  }
}
