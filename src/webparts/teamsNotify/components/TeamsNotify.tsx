import * as React from "react";
import type { ITeamsNotifyProps } from "./ITeamsNotifyProps";
import { ITeamsNotifyState } from "./ITeamsNotifyState";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
export default class TeamsNotify extends React.Component<
  ITeamsNotifyProps,
  ITeamsNotifyState
> {
  // https://www.youtube.com/watch?v=pBgcU-pAvzE
  // https://learn.microsoft.com/en-us/graph/api/userteamwork-sendactivitynotification?view=graph-rest-1.0&tabs=http
  // https://stackoverflow.com/questions/70990562/getting-error-while-using-sendactivitynotification-graph-api
  private _sendNotification(): void { 
    const endpoint: string = `https://graph.microsoft.com/v1.0/teams/62e8df43-124f-4d7c-9cd0-549984a09584`;

    const notificationBody: any = {
      topic: {
        source: "entityUrl",
        value: endpoint,
      },
      activityType: "readThisRequired",
      previewText: {
        content: "Hola Mundo",
      },
      recipient: {
        "@odata.type": "microsoft.graph.aadUserNotificationRecipient",
        userId: "54dfb99b-7ddf-44a3-b653-aa06cc639128",
      },
    };

    console.log(notificationBody);

    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((graphClient) => {
        graphClient
          .api(`${endpoint}/sendActivityNotification`)
          .post(notificationBody);
      });
  }
  private _getPeoplePickerItems(items: any[]) {
    console.log("Items:", items);
  }
  public render(): React.ReactElement<ITeamsNotifyProps> {
    const {
      // teamsContext,
      groupId,
      context,
    } = this.props;
    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory as any,
      spHttpClient: this.props.context.spHttpClient as any,
    };
    return (
      <div>
        <h1>hola mundo</h1>
        <PeoplePicker
          context={peoplePickerContext}
          titleText="People Picker"
          // webAbsoluteUrl={this.props.context.pageContext.site.absoluteUrl}
          personSelectionLimit={1}
          required={true}
          onChange={this._getPeoplePickerItems}
          ensureUser={true}
        />
        <button onClick={this._sendNotification.bind(this)}>
          Enviar notificacion
        </button>
      </div>
    );
  }
}
