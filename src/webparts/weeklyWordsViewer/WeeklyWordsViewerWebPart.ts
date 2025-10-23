import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import WeeklyWordsViewer, { IWeeklyWordsViewerProps, ViewerView, ViewerSource, ViewerDepartment } from './components/WeeklyWordsViewer';

export interface IWeeklyWordsViewerWebPartProps {
  listId: string; // set in manifest or make another property if you want
  view: ViewerView; // 'Post' | 'Tiles'
  source: ViewerSource; // 'Weekly Words' | 'Department Specific'
  department?: ViewerDepartment; // if Department Specific
  siteUrl?: string;
}

export default class WeeklyWordsViewerWebPart extends BaseClientSideWebPart<IWeeklyWordsViewerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWeeklyWordsViewerProps> = React.createElement(
      WeeklyWordsViewer,
      {
        context: this.context,
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        listId: this.properties.listId,
        view: this.properties.view || 'Tiles',
        source: this.properties.source || 'Weekly Words',
        department: this.properties.source === 'Department Specific' ? (this.properties.department || 'HR') : undefined
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

  // Re-render property pane changes immediately so the Department dropdown can enable/disable
  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, undefined, undefined);
    this.context.propertyPane.refresh();
    this.render();
  }

  private get viewOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: 'Post', text: 'Post' },
      { key: 'Tiles', text: 'Tiles' }
    ];
  }

  private get sourceOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: 'Weekly Words', text: 'Weekly Words' },
      { key: 'Department Specific', text: 'Department Specific' }
    ];
  }

  private get departmentOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: 'HR', text: 'HR' },
      { key: 'IT', text: 'IT' },
      { key: 'Marketing', text: 'Marketing' },
      { key: 'OPS', text: 'OPS' },
      { key: 'Safety', text: 'Safety' },
      { key: 'VDC', text: 'VDC' }
    ];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const departmentDisabled = this.properties.source !== 'Department Specific';

    return {
      pages: [
        {
          header: { description: 'Weekly Words Viewer' },
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('view', {
                  label: 'View',
                  options: this.viewOptions,
                  selectedKey: this.properties.view || 'Tiles'
                }),
                PropertyPaneDropdown('source', {
                  label: 'Publishing Source',
                  options: this.sourceOptions,
                  selectedKey: this.properties.source || 'Weekly Words'
                }),
                PropertyPaneDropdown('department', {
                  label: 'Department (only for Department Specific)',
                  options: this.departmentOptions,
                  disabled: departmentDisabled,
                  selectedKey: this.properties.department || 'HR'
                }),
                PropertyPaneTextField('listId', {
                  label: 'List Id'
                }),
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
