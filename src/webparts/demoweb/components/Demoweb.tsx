import * as React from 'react';
import styles from './Demoweb.module.scss';
import type { IDemowebProps } from './IDemowebProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { NewReactapp } from './NewReactapp';
export default class Demoweb extends React.Component<IDemowebProps, {}> {
  public render(): React.ReactElement<IDemowebProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <NewReactapp Context={this.props.Context}/>
    );
  }
}
