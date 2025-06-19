import * as React from 'react';
import styles from './MarketingCalendar.module.scss';
import type { IMarketingCalendarProps } from './IMarketingCalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Calendar } from '@fluentui/react';
import { GlobalLoaderProvider } from '../../../globalCommon/customLoader';
import ModernCalendar from './CalendarEvent'; // Adjust the path as needed

export default class MarketingCalendar extends React.Component<IMarketingCalendarProps> {
  public render(): React.ReactElement<IMarketingCalendarProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <GlobalLoaderProvider>
      <ModernCalendar Context={this.props.Context} CalendarTitle={this?.props?.description} MarketingCalendarId={this.props.MarketingCalendarId}/>
      </GlobalLoaderProvider>
    );
  }
}
