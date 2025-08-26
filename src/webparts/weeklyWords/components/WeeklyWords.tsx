import * as React from 'react';
import type { IWeeklyWordsProps } from './IWeeklyWordsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WeeklyWordsPost } from './WeeklyWordsPost';

export default class WeeklyWords extends React.Component<IWeeklyWordsProps> {
  public render(): React.ReactElement<IWeeklyWordsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     <WeeklyWordsPost  listId={this.props.ListId} context={this.props.context} siteUrl={this.props.siteUrl} />
    );
  }
}
