import * as React from 'react';
import { useState } from 'react';
import styles from './SharepointFeedbackForm.module.scss';
import { spfi, SPFx } from '@pnp/sp/presets/all';
import { IEmailProperties } from '@pnp/sp/sputilities';

interface IContactFormProps {
    context?: any; // SPFx context
}

interface IContactFormState {
    showForm: boolean;
    message: string;
    userEmail: string;
    userName: string;
    isSubmitting: boolean;
    submitSuccess: boolean;
    submitError: string;
}

const ContactForm: React.FC<IContactFormProps> = ({ context }) => {
    const sp = spfi().using(SPFx(context));
    const [state, setState] = useState<IContactFormState>({
        showForm: false,
        message: '',
        userEmail: '',
        userName: '',
        isSubmitting: false,
        submitSuccess: false,
        submitError: ''
    });

    // Initialize user information when component mounts
    React.useEffect(() => {
        if (context) {
            setState(prevState => ({
                ...prevState,
                userEmail: context.pageContext.user.email || '',
                userName: context.pageContext.user.displayName || ''
            }));
        }
    }, [context]);

    const handleContactClick = (): void => {
        setState(prevState => ({
            ...prevState,
            showForm: !prevState.showForm,
            submitSuccess: false,
            submitError: ''
        }));
    };

    const handleMessageChange = (event: React.ChangeEvent<HTMLTextAreaElement>): void => {
        setState(prevState => ({
            ...prevState,
            message: event.target.value
        }));
    };

    const sendEmail = async (): Promise<void> => {
        try {
            const confirmationEmailBody = `
        <html>
          <body>
            <p>Dear ${state.userName},</p>
            <p>Thank you for contacting us. Our team will review your message and get back to you.</p>
            <br/>
            <p><strong>Original Message:</strong></p>
            <div style="border-left: 3px solid #ccc; padding-left: 15px; margin-left: 10px;">
              <p>${state.message.replace(/\n/g, '<br/>')}</p>
            </div>
            <br/>
            <p>Best regards,<br/>
            Vaughn SharePoint Committee</p>
          </body>
        </html>
      `;

            const emailProps: IEmailProperties = {
                To: [state.userEmail,'vaughnsharepointcommittee@vaughnconstruction.com'],
                Subject: 'Thank you for contacting Vaughn SharePoint Support',
                Body: confirmationEmailBody,
                AdditionalHeaders: {
                    "content-type": "text/html"
                }
            };

            await sp.utility.sendEmail(emailProps);

      

            setState(prevState => ({
                ...prevState,
                isSubmitting: false,
                submitSuccess: true,
                message: '',
                showForm: false
            }));

        } catch (error) {
            console.error('Error sending email:', error);
            setState(prevState => ({
                ...prevState,
                isSubmitting: false,
                submitError: 'Failed to send message. Please try again.'
            }));
        }
    };

    const handleSubmit = async (event: React.FormEvent): Promise<void> => {
        event.preventDefault();

        if (!state.message.trim()) {
            setState(prevState => ({
                ...prevState,
                submitError: 'Please enter a message before submitting.'
            }));
            return;
        }

        setState(prevState => ({
            ...prevState,
            isSubmitting: true,
            submitError: ''
        }));

        await sendEmail();
    };

    const handleCancel = (): void => {
        setState(prevState => ({
            ...prevState,
            showForm: false,
            message: '',
            submitError: ''
        }));
    };

    return (
        <div className={styles.contactForm}>
            <div className={styles.messageContainer}>
                <p className={styles.mainMessage}>
                    If you are facing any issue or problem related to the Vaughn SharePoint or Intranet,
                    <button
                        className={styles.contactButton}
                        onClick={handleContactClick}
                        type="button"
                    >
                        Contact us
                    </button>
                </p>
            </div>

            {state.showForm && (
                <div className={styles.formContainer}>
                    <form onSubmit={handleSubmit} className={styles.contactFormElement}>
                        <div className={styles.formHeader}>
                            <h3>Contact Vaughn SharePoint Support</h3>
                        </div>

                        <div className={styles.formGroup}>
                            <label htmlFor="message" className={styles.label}>
                                Please describe your issue, suggestion, or question:
                            </label>
                            <textarea
                                id="message"
                                className={styles.textarea}
                                value={state.message}
                                onChange={handleMessageChange}
                                placeholder="Type your message here..."
                                rows={6}
                                required
                            />
                        </div>

                        {state.submitError && (
                            <div className={styles.errorMessage}>
                                {state.submitError}
                            </div>
                        )}

                        <div className={styles.buttonGroup}>
                            <button
                                type="submit"
                                className={`${styles.submitButton} ${state.isSubmitting ? styles.submitting : ''}`}
                                disabled={state.isSubmitting}
                            >
                                {state.isSubmitting ? 'Sending...' : 'Submit'}
                            </button>
                            <button
                                type="button"
                                className={styles.cancelButton}
                                onClick={handleCancel}
                                disabled={state.isSubmitting}
                            >
                                Cancel
                            </button>
                        </div>
                    </form>
                </div>
            )}

            {state.submitSuccess && (
                <div className={styles.successMessage}>
                    <p>✓ Thank you! Your message has been sent successfully. You will receive a confirmation email shortly.</p>
                </div>
            )}
        </div>
    );
};

export default ContactForm;