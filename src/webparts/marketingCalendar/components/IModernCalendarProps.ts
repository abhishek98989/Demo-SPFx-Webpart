import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IModernCalendarProps {
  description: string;
  listName: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}