// CustomDocumentSearch.tsx
import * as React from 'react';
import type { ICustomDocumentSearchProps } from './ICustomDocumentSearchProps';
import SearchDocuments from './SearchDocuments';

export default class CustomDocumentSearch extends React.Component<ICustomDocumentSearchProps> {
  public render(): React.ReactElement<ICustomDocumentSearchProps> {
    return (
      <SearchDocuments
msGraphClientFactory={this.props.msGraphClientFactory}
        context={this.props.context}
        // NEW: UI variant from property pane
        variant={this.props.searchMode || 'SearchWithResult'}
      />
    );
  }
}
