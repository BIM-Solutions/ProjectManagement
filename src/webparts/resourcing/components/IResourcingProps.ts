import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IResourcingProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  tasksListName: string;
  groupId: string;
  defaultView: "tasks" | "calendar";
  showTeamCalendar: boolean;
  context: WebPartContext;
}
