import { SPHttpClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPostsArchiveProps {
  listId: string;
  viewType: string;
  numberOfEvents: number;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  Context: any; 
  slideAfter:any;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}