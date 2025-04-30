import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IListItemCommentsProps {
  listName: string;
  itemId: number | null;  // Updated to allow null values
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}