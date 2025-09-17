import * as React from 'react';
import { useState } from 'react';
import styles from './SharepointFeedbackForm.module.scss';
import { spfi, SPFx } from '@pnp/sp/presets/all';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
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
    const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(context));
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

    const addToSharePointList = async (): Promise<void> => {
        try {
            const currentDate = new Date().toLocaleDateString();
            const listGuid = 'b1d451db-4d9e-40e7-af56-9cf204d81664';

            const listItem = {
                Title: `Feedback (${currentDate} - by ${state.userName})`,
                Feedback: state.message
            };

            await sp.web.lists.getById(listGuid).items.add(listItem);

            setState(prevState => ({
                ...prevState,
                isSubmitting: false,
                submitSuccess: true,
                message: '',
                showForm: false
            }));

        } catch (error) {
            console.error('Error adding item to SharePoint list:', error);
            setState(prevState => ({
                ...prevState,
                isSubmitting: false,
                submitError: 'Failed to submit feedback. Please try again.'
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

        await addToSharePointList();
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
                    If you are facing any issue or problem related to the Vaughn SharePoint or Intranet, &nbsp;
                    <PrimaryButton
                        text="Contact us"
                        onClick={handleContactClick}
                        type="button"
                        styles={{ root: { minWidth: '40px' } }}
                    />
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
                            <PrimaryButton
                                text={state.isSubmitting ? 'Sending...' : 'Submit'}
                                type="submit"
                                disabled={state.isSubmitting}
                            />
                            <DefaultButton
                                text="Cancel"
                                disabled={state.isSubmitting}
                                onClick={handleCancel}
                                  type="button"
                            />
                        </div>
                    </form>
                </div>
            )}

            {state.submitSuccess && (
                <div className={styles.successMessage}>
                    <p>✓ Thank you! Your feedback has been submitted successfully.</p>
                </div>
            )}
        </div>
    );
};

export default ContactForm;