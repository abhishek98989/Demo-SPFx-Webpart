import * as React from 'react';
import type { ILessonLearnedTableProps } from './ILessonLearnedTableProps';
import { escape } from '@microsoft/sp-lodash-subset';
import BlogsTable from './BlogsTable';

export default class LessonLearnedTable extends React.Component<ILessonLearnedTableProps> {
  public render(): React.ReactElement<ILessonLearnedTableProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
   <BlogsTable Context={this.props.Context} postsListId={this?.props?.postsListId}/>
    );
  }
}
