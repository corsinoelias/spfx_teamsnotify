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

    const notificationBody: any = {
      topic: {
        source: 'entityUrl',
        value: 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/9d14ae0d-9041-47d9-9b8b-e8f67d9062cd'
    },
      activityType: "approvalRequired",
      previewText: {
        content: "Hola Mundo",
      },
      templateParameters: [
        {
          name: "taskId",
          value: "Task 12322",
        },
      ],
      recipients: [
        {
            '@odata.type': 'microsoft.graph.aadUserNotificationRecipient',
            userId: '54dfb99b-7ddf-44a3-b653-aa06cc639128'
        }
      ],
      // recipient: {
      //   "@odata.type": "microsoft.graph.aadUserNotificationRecipient",
      //   userId: "admin@saldanagroup365.onmicrosoft.com",
      // },
    };

    console.log(notificationBody);

    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((graphClient) => {
        graphClient
          // .api(`${endpoint}/sendActivityNotification`)
          // .api(`/teamwork/sendActivityNotificationToRecipients`)
          // .version('beta')
          // .api(
          //   `/users/54dfb99b-7ddf-44a3-b653-aa06cc639128/teamwork/sendActivityNotification`
          // )
          .api('/teamwork/sendActivityNotificationToRecipients')
          .post(notificationBody);
        //   .get((error, response: any, rawResponse?: any) => {
        //   console.log(response)
        // });
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
