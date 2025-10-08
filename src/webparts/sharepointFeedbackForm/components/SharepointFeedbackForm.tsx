import * as React from 'react';
import type { ISharepointFeedbackFormProps } from './ISharepointFeedbackFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ContactForm from './ContactForm';
import { Scripting } from './Scripting';

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
      // <Scripting context={this.props.context} />
      <ContactForm context={this.props.context} />
    );
  }
}
