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
import WeeklyWordsViewer, { IWeeklyWordsViewerProps} from './components/WeeklyWordsViewer';
import { SPComponentLoader } from '@microsoft/sp-loader';
// Use the exact version that matches your editor build
SPComponentLoader.loadCss('https://cdn.ckeditor.com/ckeditor5/47.1.0/ckeditor5.css');
export interface IWeeklyWordsViewerWebPartProps {
  siteUrl: string;
  listId: string;
  view: 'Tiles' | 'Post' | 'Redirected'; // ⬅️ NEW
  publishingSource: 'Weekly Words' | 'Department Specific';
  department?: 'HR' | 'IT' | 'Marketing' | 'OPS' | 'Safety' | 'VDC';
}

export default class WeeklyWordsViewerWebPart extends BaseClientSideWebPart<IWeeklyWordsViewerWebPartProps> {

 // + Optional: default the view to Redirected when on Weekly-Words-Viewer.aspx
protected async onInit(): Promise<void> {
  await super.onInit();
 if (!document.getElementById('tailwind-cdn')) {
      const tailwindScript = document.createElement('script');
      tailwindScript.id = 'tailwind-cdn';
      tailwindScript.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(tailwindScript);
    }
  const url = this.context.pageContext?.site?.serverRequestPath
    || window.location.pathname
    || '';

  if (!this.properties.view) {
    if (/\/SitePages\/Weekly-Words-Viewer\.aspx/i.test(url)) {
      this.properties.view = 'Redirected';
    } else {
      this.properties.view = 'Tiles';
    }
  }

  if (!this.properties.publishingSource) {
    this.properties.publishingSource = 'Weekly Words';
  }
}


// + Ensure render passes siteUrl, listId, view, publishingSource, department correctly
public render(): void {
  const element: React.ReactElement<IWeeklyWordsViewerProps> = React.createElement(
    WeeklyWordsViewer,
    {
      context: this.context,
      siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
      listId: this.properties.listId,
      view: this.properties.view || 'Tiles',
      publishingSource: this.properties.publishingSource || 'Weekly Words',
      department:
        this.properties.publishingSource === 'Department Specific'
          ? (this.properties.department || 'HR')
          : undefined
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
protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  
  // Update the property
  if (propertyPath === 'publishingSource') {
    this.properties.publishingSource = newValue;
  }
  
  this.context.propertyPane.refresh();
  this.render();
}
  private get viewOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: 'Post', text: 'Post' },
      { key: 'Tiles', text: 'Tiles' },
      { key: 'Redirected', text: 'Redirected' }, // ⬅️ NEW
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
    const departmentDisabled = this.properties.publishingSource !== 'Department Specific';

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
                PropertyPaneDropdown('publishingSource', {
                  label: 'Publishing Source',
                  options: this.sourceOptions,
                  selectedKey: this.properties.publishingSource || 'Weekly Words'
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
