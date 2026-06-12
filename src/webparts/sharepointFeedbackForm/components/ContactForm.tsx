import * as React from 'react';
import { useState } from 'react';
import styles from './SharepointFeedbackForm.module.scss';
import { spfi, SPFx } from '@pnp/sp/presets/all';
import { DefaultButton, PrimaryButton } from '@fluentui/react';

interface IContactFormProps {
    context?: any;
}

type FormCategory = 'SharepointCommittee' | 'TechSupport' | null;

interface IContactFormState {
    activeForm: FormCategory;
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
        activeForm: null,
        message: '',
        userEmail: '',
        userName: '',
        isSubmitting: false,
        submitSuccess: false,
        submitError: ''
    });

    React.useEffect(() => {
        if (context) {
            setState(prevState => ({
                ...prevState,
                userEmail: context.pageContext.user.email || '',
                userName: context.pageContext.user.displayName || ''
            }));
        }
    }, [context]);

    const handleButtonClick = (category: FormCategory): void => {
        setState(prevState => ({
            ...prevState,
            // Toggle off if same button clicked again
            activeForm: prevState.activeForm === category ? null : category,
            message: '',
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
                Feedback: state.message,
                Category: state.activeForm   // 'SharepointCommittee' or 'TechSupport'
            };

            await sp.web.lists.getById(listGuid).items.add(listItem);

            setState(prevState => ({
                ...prevState,
                isSubmitting: false,
                submitSuccess: true,
                message: '',
                activeForm: null
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
            activeForm: null,
            message: '',
            submitError: ''
        }));
    };

    const formTitle = state.activeForm === 'TechSupport'
        ? 'Contact IT Tech Support'
        : 'Contact Vaughn SharePoint Support';

    const formPlaceholder = state.activeForm === 'TechSupport'
        ? 'Describe your IT issue here...'
        : 'Type your message here...';

    return (
        <div className={styles.contactForm}>

            {/* Row 1 — SharePoint issues */}
            <div className={styles.messageContainer}>
                <p className={styles.mainMessage}>
                    If you have questions or problems <strong>with SharePoint</strong>,&nbsp;
                    <PrimaryButton
                        text="Contact us"
                        onClick={() => handleButtonClick('SharepointCommittee')}
                        type="button"
                        styles={{ root: { minWidth: '40px' } }}
                    />
                </p>
            </div>

            {/* Row 2 — General IT issues */}
            <div className={styles.messageContainer}>
                <p className={styles.mainMessage}>
                    If you have any other IT issues,&nbsp;
                    <PrimaryButton
                        text="Email Tech Support"
                        onClick={() => handleButtonClick('TechSupport')}
                        type="button"
                        styles={{ root: { minWidth: '40px' } }}
                    />
                </p>
            </div>

            {/* Shared form — shown for either button, title/placeholder adapt to category */}
            {state.activeForm && (
                <div className={styles.formContainer}>
                    <form onSubmit={handleSubmit} className={styles.contactFormElement}>
                        <div className={styles.formHeader}>
                            <h3>{formTitle}</h3>
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
                                placeholder={formPlaceholder}
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