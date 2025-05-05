import * as React from 'react';
import styles from './LearnerBlogPost.module.scss';
import type { ILearnerBlogPostProps } from './ILearnerBlogPostProps';
import { escape } from '@microsoft/sp-lodash-subset';
import BlogPosts from './NewBlogPost';

export default class LearnerBlogPost extends React.Component<ILearnerBlogPostProps> {
  public render(): React.ReactElement<ILearnerBlogPostProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
    <BlogPosts Context={this.props.Context} PostsListId={this?.props?.postsListId} CategoriesListId={this.props?.categoriesListId}/>
    );
  }
}
