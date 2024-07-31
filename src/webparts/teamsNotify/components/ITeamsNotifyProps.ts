import { IMicrosoftTeams, WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITeamsNotifyProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  teamsContext?: IMicrosoftTeams;
  groupId: string;
  context:WebPartContext;
}
