import * as React from 'react';
import styles from './ConsolidatedCalendar.module.scss';
import type { IConsolidatedCalendarProps } from './IConsolidatedCalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { GlobalLoaderProvider } from '../../../globalCommon/customLoader';
import ConsolidatedCalendarEvent from './ConsolidatedCalendarEvent'; // Adjust the path as needed
export default class ConsolidatedCalendar extends React.Component<IConsolidatedCalendarProps> {
  public render(): React.ReactElement<IConsolidatedCalendarProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <GlobalLoaderProvider>
      <ConsolidatedCalendarEvent Context={this.props.Context} CalendarTitle={this?.props?.description} CalendarDetails={this.props.siteCalendarCombos}/>
      </GlobalLoaderProvider>
    );
  }
}
