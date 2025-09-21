import * as React from 'react';
import type { ILocationsProps } from './ILocationsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { LocationsTable } from './LocationsTable';

export default class Locations extends React.Component<ILocationsProps> {
  public render(): React.ReactElement<ILocationsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
  <LocationsTable Context={this.props.Context} />
    );
  }
}
