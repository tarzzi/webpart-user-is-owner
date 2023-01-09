import * as React from "react";
import styles from "./UserIsOwner.module.scss";
import { IUserIsOwnerProps } from "./IUserIsOwnerProps";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";
import { PrimaryButton } from "@fluentui/react";
import { Persona } from "office-ui-fabric-react/lib/Persona";
import { Stack } from "office-ui-fabric-react/lib/components/Stack";
import { DetailsList, IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { Text } from "@fluentui/react/lib/Text";
import { List } from "office-ui-fabric-react/lib/List";

interface IUser {
  id: string;
  name: string;
  mail: string;
}

interface IUserIsOwnerState {
  listItems: any[];
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
      user: {
        id: "",
        mail: "",
        name: "",
      },
    };
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
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        client
          .api(`/users/${this.state.user.mail}/ownedObjects`)
          .get((error, response) => {
            if (error) {
              console.log(error);
            }
            //filter out groups that have @odata.type === "#microsoft.graph.group"
            response.value = response.value.filter((group: any) => {
              return group["@odata.type"] === "#microsoft.graph.group";
            });
            console.log(response);
            this.setState({
              listItems: response.value,
            });
          })
          .catch((error) => {
            console.log(error);
          });
      })
      .catch((error) => {
        console.log(error);
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
  public render(): React.ReactElement<IUserIsOwnerProps> {
    const { user, listItems } = this.state;
    console.log("User state", this.state.user);
    const { hasTeamsContext, context } = this.props;
    const onRenderCell = (
      item: any
    ): JSX.Element => {
      return (
        <li key={item.id}>
          <span>{item.displayName}</span>
        </li>
      );
    };

    const stackTokens = { childrenGap: 20 };

    return (
      <section
        className={`${styles.userIsOwner} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <Stack tokens={stackTokens}>
          <Stack horizontal tokens={stackTokens}>
            <Stack.Item>
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
              {user.mail.length > 0 && ( // if user is selected
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
          <Text variant="xLarge">Group owner in:</Text>
          <List className={styles.list} items={listItems} onRenderCell={onRenderCell} />
          <PrimaryButton
            style={{ marginTop: 10 }}
            className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 margin-top-10"
            text="Get My Owned Groups"
            onClick={this._myGroups.bind(this)}
          />
        </Stack>
      </section>
    );
  }
}
