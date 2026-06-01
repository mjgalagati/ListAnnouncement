import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IListAnnouncementsProps {
  webpartTitle: string;
  sourceList: string;
  highlightTypes: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  currentUserLogin: string;
  currentUserId: number;
  isAdmin: boolean;
  itemLimit: number;
}
