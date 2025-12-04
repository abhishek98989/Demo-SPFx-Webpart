// CustomDocumentSearchWebPart.ts

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import CustomDocumentSearch from './components/CustomDocumentSearch';
import { ICustomDocumentSearchProps } from './components/ICustomDocumentSearchProps';

export interface ICustomDocumentSearchWebPartProps {
  description: string;
  searchMode: 'SearchBar' | 'SearchWithResult';
}

export default class CustomDocumentSearchWebPart extends BaseClientSideWebPart<ICustomDocumentSearchWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICustomDocumentSearchProps> = React.createElement(
      CustomDocumentSearch,
      {
        description: this.properties.description,
        context: this.context,

        // NEW: pass selected mode, fallback to SearchWithResult
        searchMode: this.properties.searchMode || 'SearchWithResult',

        // sample props (keep as in your existing file)
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
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
    const searchModeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'SearchBar', text: 'SearchBar (search box only, redirect to results page)' },
      { key: 'SearchWithResult', text: 'SearchWithResult (show results inside web part)' }
    ];

    return {
      pages: [
        {
          header: {
            description: 'Custom document search configuration'
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                // NEW DROPDOWN
                PropertyPaneDropdown('searchMode', {
                  label: 'Search mode',
                  options: searchModeOptions,
                  selectedKey: 'SearchWithResult'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // ... keep your existing theme/environment handlers etc.
}
