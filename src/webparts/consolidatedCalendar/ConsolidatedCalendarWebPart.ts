import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { TextField } from '@fluentui/react/lib/TextField';
import { IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';

import * as strings from 'ConsolidatedCalendarWebPartStrings';
import ConsolidatedCalendar from './components/ConsolidatedCalendar';
import { IConsolidatedCalendarProps } from './components/IConsolidatedCalendarProps';

// Interface for combo items
export interface IComboItem {
  siteUrl: string;
  CalendarID: string;
}

// Updated props interface
export interface IConsolidatedCalendarWebPartProps {
  description: string;
  siteCalendarCombos: IComboItem[];
  Context:any
}

// React component for the property pane field
const MultiComboField: React.FC<{
  value: IComboItem[];
  onChanged: (value: IComboItem[]) => void;
}> = ({ value, onChanged }) => {
  const [combos, setCombos] = React.useState<IComboItem[]>(
    value && value.length > 0 ? value : [{ siteUrl: '', CalendarID: '' }]
  );

  React.useEffect(() => {
    onChanged(combos);
  }, [combos, onChanged]);

  const updateCombo = (index: number, field: keyof IComboItem, newValue: string) => {
    const updatedCombos = [...combos];
    updatedCombos[index] = { ...updatedCombos[index], [field]: newValue };
    setCombos(updatedCombos);
  };

  const addCombo = () => {
    setCombos([...combos, { siteUrl: '', CalendarID: '' }]);
  };

  const removeCombo = (index: number) => {
    if (combos.length > 1) {
      const updatedCombos = combos.filter((_, i) => i !== index);
      setCombos(updatedCombos);
    }
  };

  return (
    React.createElement('div', null,
      combos.map((combo, index) => (
        React.createElement(Stack, {
          key: index,
          horizontal: true,
          tokens: { childrenGap: 8 },
          style: { marginBottom: 8 }
        },
          React.createElement(TextField, {
            label: index === 0 ? "Site URL" : "",
            value: combo.siteUrl,
            onChange: (_, newValue) => updateCombo(index, 'siteUrl', newValue || ''),
            placeholder: "Enter site URL",
            styles: { root: { width: '45%' } }
          }),
          React.createElement(TextField, {
            label: index === 0 ? "Calendar ID" : "",
            value: combo.CalendarID,
            onChange: (_, newValue) => updateCombo(index, 'CalendarID', newValue || ''),
            placeholder: "Enter calendar ID",
            styles: { root: { width: '45%' } }
          }),
          React.createElement(IconButton, {
            iconProps: { iconName: 'Delete' },
            title: "Remove",
            onClick: () => removeCombo(index),
            disabled: combos.length === 1,
            styles: { root: { marginTop: index === 0 ? 28 : 0 } }
          })
        )
      )),
      React.createElement(IconButton, {
        iconProps: { iconName: 'Add' },
        text: "Add Another",
        onClick: addCombo,
        styles: { root: { marginTop: 8 } }
      })
    )
  );
};

// Custom property pane field class
class MultiComboPropertyPaneField implements IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneCustomFieldProps;
  private _value: IComboItem[];
  private _onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private _propertiesBag: any;

  constructor(targetProperty: string, properties: {
    key: string;
    value: IComboItem[];
    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
    properties: any;
  }) {
    this.targetProperty = targetProperty;
    this._value = properties.value;
    this._onPropertyChange = properties.onPropertyChange;
    this._propertiesBag = properties.properties;

    this.properties = {
      key: properties.key,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    };
  }

  public onRender(elem: HTMLElement): void {
    const reactElement = React.createElement(MultiComboField, {
      value: this._value,
      onChanged: (value: IComboItem[]) => {
        const oldValue = this._value;
        this._value = value;
        this._onPropertyChange(this.targetProperty, oldValue, value);
        this._propertiesBag[this.targetProperty] = value; // ✅ Persist the value in the web part
      }
    });

    ReactDom.render(reactElement, elem);
  }

  public onDispose(elem: HTMLElement): void {
    ReactDom.unmountComponentAtNode(elem);
  }
}


// Factory function
const PropertyPaneMultiCombo = (targetProperty: string, properties: {
  key: string;
  value: IComboItem[];
  onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  properties: any; // Pass the web part's property bag reference
}): IPropertyPaneField<IPropertyPaneCustomFieldProps> => {
  return new MultiComboPropertyPaneField(targetProperty, {
    ...properties,
    properties: properties.properties // Ensure the required 'properties' field is included
  });
};

export default class ConsolidatedCalendarWebPart extends BaseClientSideWebPart<IConsolidatedCalendarWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IConsolidatedCalendarProps> = React.createElement(
      ConsolidatedCalendar,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        Context: this.context,
        siteCalendarCombos: this.properties.siteCalendarCombos || [] // Pass the combos to your component
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: "Calendar Settings",
              groupFields: [
                PropertyPaneMultiCombo('siteCalendarCombos', {
                  key: 'siteCalendarCombos',
                  value: this.properties.siteCalendarCombos || [],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties // Pass reference
                })
              ]
            }
          ]
        }
      ]
    };
  }
}