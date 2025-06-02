import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'PostsArchiveWebPartStrings';
import PostsArchive from './components/PostsArchive';
import { IPostsArchiveProps } from './components/IPostsArchiveProps';

export interface IPostsArchiveWebPartProps {
  listId: string;
  viewType: string;
  numberOfEvents: number;
  slideAfter:any;
}
export interface IPostItem {
  Id?: number;
  Title: string;
  PublishingDate: string;
  Source: string;
  Description: string;
  Region: string;
  Images?: string;
}
export interface IArchiveItem {
  Id: number;
  Title: string;
  PublishingDate: string;
  Source: string;
  Description: string;
  Region: string;
  Images: string;
  Attachments: any[];
  AttachmentFiles?: any[];
}

export default class PostsArchiveWebPart extends BaseClientSideWebPart<IPostsArchiveWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IPostsArchiveProps> = React.createElement(
      PostsArchive,
      {
        listId: this.properties.listId || 'Posts Archive',
        viewType: this.properties.viewType || 'table',
        numberOfEvents: this.properties.numberOfEvents || 5,
        slideAfter: this.properties.slideAfter || '3',
        spHttpClient: this.context.spHttpClient,
        Context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.listId = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listId', {
                  label: 'List Id',
                  value: this.properties.listId || 'Posts Archive'
                }),
                PropertyPaneTextField('slideAfter', {
                  label: 'Slide After (Seconds)',
                  value: this.properties.slideAfter || '3'
                }),
                PropertyPaneDropdown('viewType', {
                  label: 'View Type',
                  options: [
                    { key: 'table', text: 'Full Archive Table View' },
                    { key: 'slider', text: 'Latest Events Slider' }
                  ],
                  selectedKey: this.properties.viewType || 'table'
                }),
                PropertyPaneSlider('numberOfEvents', {
                  label: 'Number of Events (for slider view)',
                  min: 1,
                  max: 20,
                  value: this.properties.numberOfEvents || 5,
                  showValue: true,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
export interface IPostsService {
  getItems(): Promise<IArchiveItem[]>;
  getItemById(id: number): Promise<IArchiveItem>;
  createItem(item: IPostItem): Promise<void>;
  updateItem(id: number, item: Partial<IPostItem>): Promise<void>;
  deleteItem(id: number): Promise<void>;
}

// Service class for SharePoint operations
