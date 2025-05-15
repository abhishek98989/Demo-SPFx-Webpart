// AIEnabledInputField.tsx
import * as React from 'react';
import { TextField, ITextFieldProps } from '@fluentui/react/lib/TextField';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { GoogleGenAI } from "@google/genai";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { useId } from '@fluentui/react-hooks';

export interface IAIEnabledInputFieldProps extends ITextFieldProps {
    perplexityApiKey: string;
}

const styles = mergeStyleSets({
    inputWrapper: {
        position: 'relative',
    },
    actionsContainer: {
        position: 'absolute',
        top: 0,
        right: 0,
        padding: '5px',
        display: 'flex',
        gap: '5px',
        opacity: 0,
        transition: 'opacity 0.3s ease',
    },
    visible: {
        opacity: 1,
    },
    dialogContent: {
        maxWidth: '600px',
        minWidth: '320px',
    },
});

export const AIEnabledInputField: React.FunctionComponent<IAIEnabledInputFieldProps> = (props) => {
    const { perplexityApiKey, ...textFieldProps } = props;

    const [value, setValue] = React.useState(props.defaultValue || props.value || '');
    const [isHovered, setIsHovered] = React.useState(false);
    const [isDialogOpen, setIsDialogOpen] = React.useState(false);
    const [aiSuggestion, setAiSuggestion] = React.useState('');
    const [isLoading, setIsLoading] = React.useState(false);
    const [dialogType, setDialogType] = React.useState<'rewrite' | 'grammar'>('rewrite');

    const tooltipIdRewrite = useId('tooltip-rewrite');
    const tooltipIdGrammar = useId('tooltip-grammar');
     const ai = new GoogleGenAI({ apiKey: "AIzaSyDKDaZsYAgCuSwle7XWoVbhb3xnlYZ0vvo" });
    // const ai = new GoogleGenerativeAI( "AIzaSyDKDaZsYAgCuSwle7XWoVbhb3xnlYZ0vvo");
    React.useEffect(() => {
        if (props.value !== undefined && props.value !== value) {
            setValue(props.value);
        }
    }, [props.value]);

    const handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        setValue(newValue || '');
        if (props.onChange) {
            props.onChange(event, newValue);
        }
    };

    const handleMouseEnter = () => {
        setIsHovered(true);
    };

    const handleMouseLeave = () => {
        setIsHovered(false);
    };
 
    async function generateAiContent(type: 'rewrite' | 'grammar' , textValue:any) {
        if (!value.trim()) return;
        setIsLoading(true);
        setDialogType(type);
      
        try {
          // Construct the appropriate prompt
          const prompt = 
            type === 'rewrite'
              ? `Rewrite the following sentence using clearer, more natural wording. Return only one improved version and no explanation: "${textValue}"`
              : `Fix grammar errors in the following text without changing the meaning. Return only one improved version and no explanation:  "${textValue}"`;
      
          // Call the new GenAI model
          const response = await ai.models.generateContent({
            model: "gemini-2.0-flash",
            contents: prompt,
          });
      
          // The SDK returns the suggestion in `text`
          const aiResponse:any = response.text;
      
          // Update UI with the AI suggestion
          setAiSuggestion(aiResponse);
          setIsDialogOpen(true);
        } catch (error) {
          console.error('Error calling Google GenAI API:', error);
          alert('Failed to process with AI. Please try again later.');
        } finally {
          setIsLoading(false);
        }
      }
    const handleAiRewrite = () => {
        generateAiContent('rewrite',value);
    };

    const handleGrammarFix = () => {
        generateAiContent('grammar',value);
    };

    const handleAccept = () => {
        setValue(aiSuggestion);
        if (props.onChange) {
            // Creating a synthetic event to pass to onChange
            const syntheticEvent = {
                target: { value: aiSuggestion }
            } as unknown as React.FormEvent<HTMLInputElement | HTMLTextAreaElement>;
            props.onChange(syntheticEvent, aiSuggestion);
        }
        setIsDialogOpen(false);
    };

    const handleRegenerate = () => {
        generateAiContent('rewrite',aiSuggestion);
    };

    const handleCancel = () => {
        setIsDialogOpen(false);
    };

    const dialogTitle = dialogType === 'rewrite' ? 'AI Rewrite Suggestion' : 'Grammar Correction';

    return (
        <div
            className={styles.inputWrapper}
            onMouseEnter={handleMouseEnter}
            onMouseLeave={handleMouseLeave}
        >
            <TextField
                {...textFieldProps}
                value={value}
                onChange={handleInputChange}
            />

            <div className={`${styles.actionsContainer} ${isHovered ? styles.visible : ''}`}>
                <TooltipHost
                    content="Fix Grammar"
                    id={tooltipIdGrammar}
                >
                    <IconButton
                        iconProps={{ iconName: 'Spelling' }}
                        title="Fix Grammar"
                        ariaLabel="Fix Grammar"
                        onClick={handleGrammarFix}
                        disabled={isLoading || !value.trim()}
                    />
                </TooltipHost>
                <TooltipHost
                    content="AI Rewrite"
                    id={tooltipIdRewrite}
                >
                    <IconButton
                        iconProps={{ iconName: 'EditCreate' }}
                        title="AI Rewrite"
                        ariaLabel="AI Rewrite"
                        onClick={handleAiRewrite}
                        disabled={isLoading || !value.trim()}
                    />
                </TooltipHost>
            </div>

            {isLoading && (
                <Spinner
                    size={SpinnerSize.medium}
                    label="Processing with AI..."
                    styles={{ root: { position: 'absolute', top: '50%', right: '60px', transform: 'translateY(-50%)' } }}
                />
            )}

            <Dialog
                hidden={!isDialogOpen}
                onDismiss={handleCancel}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    subText: dialogType === 'rewrite'
                        ? 'Here is the AI-generated rewrite of your text:'
                        : 'Here is the grammar-corrected version of your text:',
                    className: styles.dialogContent
                }}
                modalProps={{ isBlocking: true }}
            >
                <TextField
                    multiline
                    rows={6}
                    value={aiSuggestion}
                    readOnly
                />
                <DialogFooter>
                    <PrimaryButton onClick={handleAccept} text="Accept" />
                    {dialogType === 'rewrite' &&<DefaultButton onClick={handleRegenerate} text="Regenerate" /> }
                    <DefaultButton onClick={handleCancel} text="Cancel" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};