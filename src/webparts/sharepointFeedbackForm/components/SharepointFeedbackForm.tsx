import * as React from 'react';
import type { ISharepointFeedbackFormProps } from './ISharepointFeedbackFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ContactForm from './ContactForm';

export default class SharepointFeedbackForm extends React.Component<ISharepointFeedbackFormProps> {
  public render(): React.ReactElement<ISharepointFeedbackFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <ContactForm context={this.props.context} />
    );
  }
}
