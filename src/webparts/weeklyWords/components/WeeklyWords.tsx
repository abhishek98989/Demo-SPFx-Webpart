import * as React from 'react';
import type { IWeeklyWordsProps } from './IWeeklyWordsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WeeklyWordsPost } from './WeeklyWordsPost';
import { SPComponentLoader } from '@microsoft/sp-loader';
export default class WeeklyWords extends React.Component<IWeeklyWordsProps> {
  public render(): React.ReactElement<IWeeklyWordsProps> {
     SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/scss/vendor/_rfs.scss");
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
