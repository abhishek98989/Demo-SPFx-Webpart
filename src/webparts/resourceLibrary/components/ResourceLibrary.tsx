import * as React from 'react';
import type { IResourceLibraryProps } from './IResourceLibraryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import DocumentLibrary from './DocumentLibrary';

export default class ResourceLibrary extends React.Component<IResourceLibraryProps> {
  public render(): React.ReactElement<IResourceLibraryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     <DocumentLibrary context={this.props.context} />
    );
  }
}
