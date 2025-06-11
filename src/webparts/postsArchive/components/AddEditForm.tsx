// AddEditForm.tsx
import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Panel,
  PanelType,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Label,
  Stack,
  IconButton,
  Text
} from '@fluentui/react';
import {
  useToastController,
  Toaster,
  ToastTitle,
  Toast,
  ToastBody,
} from '@fluentui/react-components';
import 'bootstrap/dist/css/bootstrap.min.css';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { SPHttpClient } from '@microsoft/sp-http';
import { IArchiveItem } from '../PostsArchiveWebPart';
import styles from './PostsArchive.module.scss';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import moment from 'moment';
import './PostsStyle.css';
import AttachmentViewer from './AttachmentViewer';
import { useGlobalLoaderContext } from '../../../globalCommon/customLoader';

export interface IPostItem {
  Id?: number;
  Title: string;
  PublishingDate: string;
  Source: string;
  Description: string;
  Region: string;
  CreatedBy?: string;
  ModifiedBy?: string;
  Created?: string;
  Modified?: string;
  AuthorId?: number;
  EditorId?: number;
}

export interface IAddEditFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSave: () => void;
  item?: any;
  listId: string;
  Context: any;
}

const AddEditForm: React.FC<IAddEditFormProps> = (props) => {
  const { showLoader, hideLoader } = useGlobalLoaderContext();
  const [formData, setFormData] = useState<IPostItem>({
    Title: '',
    PublishingDate: new Date().toISOString().split('T')[0],
    Source: '',
    Description: '',
    Region: ''
  });
  
  const sp = spfi().using(SPFx(props?.Context));
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [attachments, setAttachments] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<any[]>([]);
  
  // Toast controller
  const { dispatchToast } = useToastController();
  
  // File viewer states
  const [isViewerOpen, setIsViewerOpen] = useState<boolean>(false);
  const [currentViewFile, setCurrentViewFile] = useState<{ url: string; name: string; type: string } | null>(null);
  
  const regionOptions: IDropdownOption[] = [
    { key: 'Houston', text: 'Houston' },
    { key: 'Austin', text: 'Austin' },
    { key: 'Lubbock', text: 'Lubbock' },
    { key: 'El Paso', text: 'El Paso' },
    { key: 'DFW', text: 'DFW' },
    { key: 'College Station', text: 'College Station' },
    { key: 'RGV', text: 'RGV' },
    { key: 'San Antonio', text: 'San Antonio' },
    { key: 'BCS/Stephenville', text: 'BCS/Stephenville' },
    { key: 'North Texas', text: 'North Texas' },
    { key: 'Galveston', text: 'Galveston' },
    { key: 'Louisiana', text: 'Louisiana' },
    { key: 'Rio Grande Valley', text: 'Rio Grande Valley' }
  ];

  // Toast notification helper
  const showToast = (title: string, message: string, intent: 'success' | 'error' | 'warning' | 'info' = 'info') => {
    dispatchToast(
      <Toast>
        <ToastTitle>{title}</ToastTitle>
        <ToastBody>{message}</ToastBody>
      </Toast>,
      { intent, timeout: 4000 }
    );
  };

  useEffect(() => {
    if (props.item) {
      
      // Load existing attachments if editing
      if (props.item.Id) {
        const loadAndSetFormData = async () => {
            let Item = await getItemById(props.item.Id);
            setFormData({
              ...Item,
              PublishingDate: Item.PublishingDate ? 
                new Date(Item.PublishingDate).toISOString().split('T')[0] : 
                new Date().toISOString().split('T')[0]
            });
            
        }
        loadAndSetFormData();
        loadExistingAttachments(props.item.Id);
      }
    } else {
      resetForm();
    }
  }, [props.item, props.isOpen]);

  const loadExistingAttachments = async (itemId: number): Promise<void> => {
    try {
      const item = await sp.web.lists
        .getById(props.listId)
        .items
        .getById(itemId)
        .select("AttachmentFiles")
        .expand("AttachmentFiles")();
      
      setExistingAttachments(item.AttachmentFiles || []);
    } catch (error) {
      console.error('Error loading attachments:', error);
    }
  };

  const resetForm = (): void => {
    setFormData({
      Title: '',
      PublishingDate: new Date().toISOString().split('T')[0],
      Source: '',
      Description: '',
      Region: ''
    });
    setErrors({});
    setAttachments([]);
    setExistingAttachments([]);
  };

  const validateForm = (): boolean => {
    const newErrors: { [key: string]: string } = {};

    if (!formData.Title.trim()) {
      newErrors.Title = 'Title is required';
    }

    if (!formData.PublishingDate) {
      newErrors.PublishingDate = 'Publishing Date is required';
    }

    if (!formData.Source.trim()) {
      newErrors.Source = 'Source is required';
    }
    if (!formData.Region) {
      newErrors.Region = 'Region is required';
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleInputChange = (field: keyof IPostItem, value: any): void => {
    setFormData((prev: any) => ({
      ...prev,
      [field]: value
    }));

    // Clear error for this field
    if (errors[field]) {
      setErrors(prev => ({
        ...prev,
        [field]: ''
      }));
    }
  };

  const handleRichTextChange = (text: string): string => {
    handleInputChange('Description', text);
    return text;
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files) {
      const fileArray = Array.from(files);
      // Filter for images and PDFs only
      const allowedFiles = fileArray.filter(file => {
        const fileType = file.type;
        return fileType.startsWith('image/') || fileType === 'application/pdf';
      });
      
      if (allowedFiles.length !== fileArray.length) {
        showToast('File Type Warning', 'Only image files and PDFs are allowed.', 'warning');
      }
      
      setAttachments(prev => [...prev, ...allowedFiles]);
    }
  };

  const removeAttachment = (index: number): void => {
    setAttachments(prev => prev.filter((_, i) => i !== index));
  };

  const removeExistingAttachment = async (fileName: string): Promise<void> => {
    if (!props.item?.Id) return;
    
    try {
      await sp.web.lists
        .getById(props.listId)
        .items
        .getById(props.item.Id)
        .attachmentFiles
        .getByName(fileName)
        .delete();
      
      setExistingAttachments(prev => prev.filter(att => att.FileName !== fileName));
      showToast('Success', 'Attachment removed successfully!', 'success');
    } catch (error) {
      console.error('Error removing attachment:', error);
      showToast('Error', 'Error removing attachment.', 'error');
    }
  };

  // File viewer functions
  const viewExistingFile = async (attachment: any): Promise<void> => {
    try {
      const fileUrl = `${props.Context.pageContext.web.absoluteUrl}/_api/web/lists('${props.listId}')/items(${props.item.Id})/AttachmentFiles('${attachment.FileName}')/$value`;
      const fileExtension = attachment.FileName.split('.').pop()?.toLowerCase();
      const fileType = fileExtension === 'pdf' ? 'application/pdf' : 'image';
      
      setCurrentViewFile({
        url: fileUrl,
        name: attachment.FileName,
        type: fileType
      });
      setIsViewerOpen(true);
    } catch (error) {
      console.error('Error viewing file:', error);
      showToast('Error', 'Error opening file for viewing.', 'error');
    }
  };

  const viewNewFile = (file: File): void => {
    const fileUrl = URL.createObjectURL(file);
    const fileType = file.type.startsWith('image/') ? 'image' : 'application/pdf';
    
    setCurrentViewFile({
      url: fileUrl,
      name: file.name,
      type: fileType
    });
    setIsViewerOpen(true);
  };

  const closeViewer = (): void => {
    setIsViewerOpen(false);
    if (currentViewFile?.url.startsWith('blob:')) {
      URL.revokeObjectURL(currentViewFile.url);
    }
    setCurrentViewFile(null);
  };

  const DeleteButton = ({ onDelete }:any) => (
    <IconButton
      iconProps={{ iconName: 'Delete' }}
      title="Delete"
      ariaLabel="Delete"
      onClick={onDelete}
      styles={{ root: { color: 'red' } }}
    />
  );

  const getFileIcon = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase();
    if (extension === 'pdf') {
      return 'PDF';
    } else if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg'].includes(extension || '')) {
      return 'FileImage';
    }
    return 'Attach';
  };

  const handleSave = async (): Promise<void> => {
    if (!validateForm()) {
      showToast('Validation Error', 'Please fix the validation errors before saving.', 'error');
      return;
    }
    showLoader()
    setIsLoading(true);

    try {
      let itemId: number;
      
      if (props.item?.Id) {
        // Update existing item
        await updateItem(props.item.Id, formData);
        itemId = props.item.Id;
        hideLoader()
        showToast('Success', 'Post updated successfully!', 'success');
      } else {
        // Create new item
        const result = await createItem(formData);
        itemId = result.Id;
        hideLoader()
        showToast('Success', 'Post created successfully!', 'success');
      }

      // Handle file attachments if any
      if (attachments.length > 0) {
        await uploadAttachments(itemId);
      }

      setTimeout(() => {
        props.onSave();
        props.onDismiss();
      }, 500);

    } catch (error) {
      console.error('Error saving post:', error);
      showToast('Error', 'An error occurred while saving the post. Please try again.', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const uploadAttachments = async (itemId: number): Promise<void> => {
    try {
      for (const file of attachments) {
        const arrayBuffer = await file.arrayBuffer();
        await sp.web.lists
          .getById(props.listId)
          .items
          .getById(itemId)
          .attachmentFiles
          .add(file.name, arrayBuffer);
      }
      console.log('All attachments uploaded successfully');
    } catch (error) {
      console.error('Error uploading attachments:', error);
      throw error;
    }
  };

  async function getItems(): Promise<IArchiveItem[]> {
    try {
      const items = await sp.web.lists
        .getById(props.listId)
        .items
        .select(
          "Id", "Title", "PublishingDate", "Source", "Description", "Region", 
          "Attachments", "AttachmentFiles", "Created", "Modified",
          "Author/Title", "Editor/Title", "Author/EMail", "Editor/EMail"
        )
        .expand("AttachmentFiles", "Author", "Editor")
        .orderBy("PublishingDate", false)();

      return items.map(item => ({
        ...item,
        CreatedBy: item.Author?.Title || 'Unknown',
        ModifiedBy: item.Editor?.Title || 'Unknown',
        CreatedByEmail: item.Author?.EMail || '',
        ModifiedByEmail: item.Editor?.EMail || ''
      })) as IArchiveItem[];
    } catch (error) {
      console.error("Error fetching items:", error);
      throw error;
    }
  }

  async function getItemById(id: number): Promise<IArchiveItem> {
    try {
      const item = await sp.web.lists
        .getById(props.listId)
        .items
        .getById(id)
        .select(
          "Id", "Title", "PublishingDate", "Source", "Description", "Region", 
          "Attachments", "AttachmentFiles", "Created", "Modified",
          "Author/Title", "Editor/Title", "Author/EMail", "Editor/EMail"
        )
        .expand("AttachmentFiles", "Author", "Editor")();

      return {
        ...item,
        CreatedBy: item.Author?.Title || 'Unknown',
        ModifiedBy: item.Editor?.Title || 'Unknown',
        CreatedByEmail: item.Author?.EMail || '',
        ModifiedByEmail: item.Editor?.EMail || ''
      } as IArchiveItem;
    } catch (error) {
      console.error("Error fetching item:", error);
      throw error;
    }
  }

  async function createItem(item: IPostItem): Promise<any> {
    try {
      const postItem = {
        Title: item.Title,
        PublishingDate: moment(item.PublishingDate, 'YYYY-MM-DD').format('YYYY-MM-DDTHH:mm:ss') + 'Z',
        Source: item.Source,
        Description: item.Description,
        Region: item.Region
      };
      
      const result = await sp.web.lists
        .getById(props.listId)
        .items
        .add(postItem);
      
      return result;
    } catch (error) {
      console.error("Error creating item:", error);
      throw error;
    }
  }

  async function updateItem(id: number, item: Partial<IPostItem>): Promise<void> {
    try {
      const updateData = {
        Title: item.Title,
        PublishingDate: item.PublishingDate ? 
          moment(item.PublishingDate, 'YYYY-MM-DD').format('YYYY-MM-DDTHH:mm:ss') + 'Z' : 
          undefined,
        Source: item.Source,
        Description: item.Description,
        Region: item.Region
      };

      await sp.web.lists
        .getById(props.listId)
        .items
        .getById(id)
        .update(updateData);
    } catch (error) {
      console.error("Error updating item:", error);
      throw error;
    }
  }

  async function deleteItem(id: number): Promise<void> {
    try {
      if(confirm('Are you sure you want to delete this post? This action cannot be undone.')) {
        await sp.web.lists
          .getById(props.listId)
          .items
          .getById(id)
          .recycle();
        
        showToast('Success', 'Post deleted successfully!', 'success');
        
        setTimeout(() => {
          props.onSave();
          props.onDismiss();
        }, 1500);
      }
    } catch (error) {
      console.error("Error deleting item:", error);
      showToast('Error', 'Error deleting post. Please try again.', 'error');
      throw error;
    }
  }

  const handleCancel = (): void => {
    resetForm();
    props.onDismiss();
  };

  const renderUserInfo = (): JSX.Element | null => {
    if (!props.item || !props.item.Id) return null;

    return (
      <Stack tokens={{ childrenGap: 10 }} className={styles.userInfoSection}>
        <Stack horizontal tokens={{ childrenGap: 20 }}>
          <Stack tokens={{ childrenGap: 5 }}>
            <Text variant="small">
              <strong>Created by:</strong> {formData?.CreatedBy || 'Unknown'}
            </Text>
            <Text variant="small">
              <strong>Created on:</strong> {formData?.Created ? 
                moment(formData.Created).format('MMM DD, YYYY hh:mm A') : 'Unknown'}
            </Text>
          </Stack>
          <Stack tokens={{ childrenGap: 5 }}>
            <Text variant="small">
              <strong>Last modified by:</strong> {formData.ModifiedBy || 'Unknown'}
            </Text>
            <Text variant="small">
              <strong>Modified on:</strong> {formData.Modified ? 
                moment(formData.Modified).format('MMM DD, YYYY hh:mm A') : 'Unknown'}
            </Text>
          </Stack>
        </Stack>
      </Stack>
    );
  };

  const renderFooterContent = (): JSX.Element => (
    <div className='d-flex footerInfo justify-content-between'>
        {renderUserInfo()}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
        
     {props?.item?.Id &&
     <DefaultButton
    text="Delete"
    disabled={isLoading}
    styles={{
      root: {
        backgroundColor: '#e0e0e0', // light grey
        color: '#000',
        border: '1px solid #c8c8c8',
      },
      rootHovered: {
        backgroundColor: '#d0d0d0',
      },
      rootPressed: {
        backgroundColor: '#c0c0c0',
      },
    }}
    onClick={() => deleteItem(props?.item?.Id)}
  />}
      <PrimaryButton
        text={props.item ? 'Update' : 'Save'}
        onClick={handleSave}
        disabled={isLoading}
      />
      <DefaultButton
        text="Cancel"
        onClick={handleCancel}
        disabled={isLoading}
      />
    </Stack>
    </div>

  );

  return (
    <>
      <Toaster />
      <Panel
        isOpen={props.isOpen}
        onDismiss={props.onDismiss}
        type={PanelType.large}
        headerText={props.item ? 'Edit Post' : 'Add New Post'}
        onRenderFooterContent={renderFooterContent}
        isFooterAtBottom={true}
        isBlocking={false}
        className={styles.addEditPanel}
      >
        <div className={styles.formContainer}>
          <Stack tokens={{ childrenGap: 20 }}>
            <TextField
              label="Title"
              value={formData.Title}
              onChange={(_, value) => handleInputChange('Title', value || '')}
              errorMessage={errors.Title}
              required
            />

            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <TextField
                label="Publishing Date"
                type="date"
                value={formData.PublishingDate}
                onChange={(_, value) => handleInputChange('PublishingDate', value || '')}
                errorMessage={errors.PublishingDate}
                required
                style={{ flex: 1 }}
              />

              <Dropdown
                label="Region"
                options={regionOptions}
                selectedKey={formData.Region}
                onChange={(_, option) => handleInputChange('Region', option?.key || '')}
                errorMessage={errors.Region}
                required
                style={{ flex: 1 , width: '300px' }}
              />
            </Stack>

            <TextField
              label="Source"
              value={formData.Source}
              onChange={(_, value) => handleInputChange('Source', value || '')}
              errorMessage={errors.Source}
              required
            />

            <div className={styles.richTextSection}>
              <Label>Description</Label>
              <RichText
                value={formData.Description}
                onChange={handleRichTextChange}
                isEditMode={true}
                className={styles.richTextEditor}
              />
              {errors.Description && (
                <Text variant="small" className={styles.errorText}>
                  {errors.Description}
                </Text>
              )}
            </div>

            <div className={styles.attachmentSection}>
              <Label>Attachments (Images only)</Label>
              
              {/* File Input - Always on top */}
              <div className={styles.fileInputSection}>
                <input
                  type="file"
                  multiple
                  accept="image/*"
                  onChange={handleFileChange}
                  className={styles.fileInput}
                />
              </div>

              {/* New Attachments */}
              {attachments.length > 0 && (
                <div className={styles.attachmentList}>
                  <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>New Attachments:</Text>
                  {attachments.map((file, index) => (
                    <div key={index} className={styles.attachmentItem}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <IconButton
                          iconProps={{ iconName: getFileIcon(file.name) }}
                          onClick={() => viewNewFile(file)}
                          title={`View ${file.name}`}
                          className={styles.viewButton}
                        />
                        <span 
                          style={{ cursor: 'pointer', flex: 1 }}
                          onClick={() => viewNewFile(file)}
                          title="Click to view"
                        >
                          {file.name} ({(file.size / 1024 / 1024).toFixed(2)} MB)
                        </span>
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          onClick={() => removeAttachment(index)}
                          className={styles.removeButton}
                          title="Remove attachment"
                        />
                      </Stack>
                    </div>
                  ))}
                </div>
              )}

              {/* Existing Attachments */}
              {existingAttachments.length > 0 && (
                <div className={styles.existingAttachments}>
                  <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>Existing Attachments:</Text>
                  {existingAttachments.map((attachment, index) => (
                    <div key={index} className={styles.attachmentItem}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <IconButton
                          iconProps={{ iconName: getFileIcon(attachment.FileName) }}
                          onClick={() => viewExistingFile(attachment)}
                          title={`View ${attachment.FileName}`}
                          className={styles.viewButton}
                        />
                        <span 
                          style={{ cursor: 'pointer', flex: 1 }}
                          onClick={() => viewExistingFile(attachment)}
                          title="Click to view"
                        >
                          {attachment.FileName}
                        </span>
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          onClick={() => removeExistingAttachment(attachment.FileName)}
                          className={styles.removeButton}
                          title="Remove attachment"
                        />
                      </Stack>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </Stack>
        </div>
      </Panel>
      <AttachmentViewer isOpen={isViewerOpen} currentFile={currentViewFile} onClose={closeViewer}/>
    </>
  );
};

export default AddEditForm;