import * as React from 'react';
import type { IWeeklyWordsProps } from './IWeeklyWordsProps';
import { WeeklyWordsPost } from './WeeklyWordsPost';
import { WeeklyWordsArchive } from './WeeklyWordsArchive';
import { SPComponentLoader } from '@microsoft/sp-loader';

export default class WeeklyWords extends React.Component<IWeeklyWordsProps> {
  public render(): React.ReactElement<IWeeklyWordsProps> {
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/scss/vendor/_rfs.scss");

    const { viewMode } = this.props;

    if (viewMode === 'archive') {
      return (
        <WeeklyWordsArchive
          listId={this.props.ListId}
          context={this.props.context}
          siteUrl={this.props.siteUrl}
        />
      );
    }

    return (
      <WeeklyWordsPost
        listId={this.props.ListId}
        context={this.props.context}
        siteUrl={this.props.siteUrl}
        title={this.props.title}
      />
    );
  }
}
