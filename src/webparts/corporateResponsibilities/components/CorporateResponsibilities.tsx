import * as React from 'react';
import type { ICorporateResponsibilitiesProps } from './ICorporateResponsibilitiesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import CorporateResponsibilitiesTable from './CorporateResponsibilitiesTable';

export default class CorporateResponsibilities extends React.Component<ICorporateResponsibilitiesProps> {
  public render(): React.ReactElement<ICorporateResponsibilitiesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
  return (
    // Make sure CorporateResponsibilitiesTable is imported or defined
    <CorporateResponsibilitiesTable Context={this.props.Context} />
  );
  }
}
