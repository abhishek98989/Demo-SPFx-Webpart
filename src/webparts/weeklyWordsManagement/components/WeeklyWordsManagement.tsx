import * as React from 'react';
import styles from './WeeklyWordsManagement.module.scss';
import type { IWeeklyWordsManagementProps } from './IWeeklyWordsManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import NewsManager from './NewsManagement';

export default class WeeklyWordsManagement extends React.Component<IWeeklyWordsManagementProps> {
  public render(): React.ReactElement<IWeeklyWordsManagementProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
   <NewsManager context={this.props?.context} listId={this.props?.listId} siteUrl={this?.props?.siteUrl}/>
    );
  }
}
