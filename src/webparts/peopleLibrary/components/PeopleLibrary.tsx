import * as React from 'react';
import styles from './PeopleLibrary.module.scss';
import type { IPeopleLibraryProps } from './IPeopleLibraryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import UsersTable from './UsersTable';

export default class PeopleLibrary extends React.Component<IPeopleLibraryProps> {
  public render(): React.ReactElement<IPeopleLibraryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <UsersTable context={this.props.context} />
    );
  }
}
