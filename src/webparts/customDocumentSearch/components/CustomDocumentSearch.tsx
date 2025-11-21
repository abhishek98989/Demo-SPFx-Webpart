import * as React from 'react';
import type { ICustomDocumentSearchProps } from './ICustomDocumentSearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import SearchDocuments from './SearchDocuments';

export default class CustomDocumentSearch extends React.Component<ICustomDocumentSearchProps> {
  public render(): React.ReactElement<ICustomDocumentSearchProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
    <SearchDocuments context={this.props?.context} />
    );
  }
}
