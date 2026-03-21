import React from 'react';
import { useState, useEffect } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import {
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  IStackStyles,
  Label,
  MessageBar,
  MessageBarType,
  DayOfWeek,
  Text,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { TimePicker } from '@fluentui/react/lib/TimePicker';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Breadcrumb, IBreadcrumbItem } from '@fluentui/react/lib/Breadcrumb';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { EventRecurrenceInfo } from '../../../globalCommon/EventRecurrenceControls/EventRecurrenceInfo/EventRecurrenceInfo';
import { spfi, SPFx } from '@pnp/sp';
import { IAttachmentInfo } from '@pnp/sp/attachments';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import { parseRecurrenceToString } from '../../../globalCommon/reccurenceStringToText';
import moment from 'moment';

// ─── Hardcoded SP paths ────────────────────────────────────────────────────────
const PDF_SP_ROOT = 'https://vaughnconstruction.sharepoint.com';
const PDF_FOLDER_SERVER_RELATIVE = '/Forms/Memos';

const IMG_SP_ROOT = 'https://vaughnconstruction.sharepoint.com/news';
const IMG_FOLDER_SERVER_RELATIVE = '/news/PublishingImages';

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg']);

const stackTokens: IStackTokens = { childrenGap: 12 };
const stackStyles: IStackStyles = { root: { padding: '0 0 12px 0' } };

// ─── Helpers ──────────────────────────────────────────────────────────────────
function parseError(e: any): string {
  if (!e) return 'Unknown error';
  if (typeof e === 'string') return e;
  if (e?.message) return e.message;
  try { return JSON.stringify(e); } catch { return String(e); }
}

function getFileExt(name: string): string {
  const m = name.match(/(\.[^.]+)$/);
  return m ? m[1].toLowerCase() : '';
}

function buildBreadcrumb(
  currentPath: string,
  navigate: (path: string) => void
): IBreadcrumbItem[] {
  const parts = currentPath.replace(/^\//, '').split('/');
  return parts.map((seg, idx) => {
    const path = '/' + parts.slice(0, idx + 1).join('/');
    return { key: path, text: seg, onClick: () => navigate(path) };
  });
}

// ─── View-Only Dialog ─────────────────────────────────────────────────────────
// AFTER
const EventViewDialog: React.FC<{
  event: any;
  attachments: IAttachmentInfo[];
  onClose: () => void;
  categoryBg?: string;
  categoryFg?: string;
}> = ({ event, attachments, onClose, categoryBg = '#fff4ce', categoryFg = '#7a5c00' }) => {
  const formatDate = (d: Date) => moment(d).format('MMMM D, YYYY h:mm A');
  const pdfAttachments = attachments.filter(a => a.FileName.toLowerCase().endsWith('.pdf'));

  return (
    <Dialog
      hidden={false}
      onDismiss={onClose}
      dialogContentProps={{
        type: DialogType.normal,
        title: event.title || 'Event Details',
        styles: {
          title: {
            fontSize: 22,
            fontWeight: 700,
            paddingBottom: 12,
            borderBottom: '2px solid #f3f2f1',
            marginBottom: 4,
          },
          inner: { padding: '0 24px 4px 24px' },
        },
      }}
      modalProps={{
        isBlocking: false,
        styles: {
          main: {
            borderRadius: 8,
            boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
            overflow: 'hidden',
          },
        },
      }}
      minWidth={660}
      maxWidth={760}
    >
      {/* ── Date / Time Block ── */}
      <Stack
        horizontal
        tokens={{ childrenGap: 0 }}
        styles={{
          root: {
            background: '#f8f7f6',
            border: '1px solid #edebe9',
            borderRadius: 6,
            marginBottom: 16,
            overflow: 'hidden',
          },
        }}
      >
        <Stack
          tokens={{ childrenGap: 3 }}
          styles={{ root: { flex: 1, padding: '14px 20px', borderRight: '1px solid #edebe9' } }}
        >
          <Text
            variant="tiny"
            styles={{
              root: { color: '#8a8886', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 },
            }}
          >
            Start
          </Text>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#201f1e' } }}>
            {event.fAllDayEvent
              ? moment(event.startTime).format('MMMM D, YYYY')
              : formatDate(event.startTime)}
          </Text>
          {event.fAllDayEvent && (
            <Text variant="small" styles={{ root: { color: '#8a8886' } }}>All Day</Text>
          )}
        </Stack>

        <Stack tokens={{ childrenGap: 3 }} styles={{ root: { flex: 1, padding: '14px 20px' } }}>
          <Text
            variant="tiny"
            styles={{
              root: { color: '#8a8886', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 },
            }}
          >
            End
          </Text>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#201f1e' } }}>
            {event.fAllDayEvent
              ? moment(event.endTime).format('MMMM D, YYYY')
              : formatDate(event.endTime)}
          </Text>
          {event.fAllDayEvent && (
            <Text variant="small" styles={{ root: { color: '#8a8886' } }}>All Day</Text>
          )}
        </Stack>
      </Stack>

      {/* ── Location + Category Pills ── */}
      {(event.locations || event.category) && (
        <Stack
          horizontal
          tokens={{ childrenGap: 8 }}
          styles={{ root: { marginBottom: 16, flexWrap: 'wrap' } }}
        >
          {event.locations && (
            <Stack
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 6 }}
              styles={{
                root: {
                  background: '#fff',
                  border: '1px solid #edebe9',
                  borderRadius: 20,
                  padding: '5px 14px',
                  boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
                },
              }}
            >
              <span style={{ fontSize: 14 }}>📍</span>
              <Text variant="small" styles={{ root: { color: '#323130', fontWeight: 500 } }}>
                {event.locations}
              </Text>
            </Stack>
          )}
        
          {event.category && (
            <Stack
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 6 }}
              styles={{
                root: {
                  background: categoryBg,
                  border: `1px solid ${categoryBg}cc`,
                  borderRadius: 20,
                  padding: '5px 14px',
                },
              }}
            >
              <span style={{ fontSize: 14 }}>🏷️</span>
              <Text variant="small" styles={{ root: { color: categoryFg, fontWeight: 500 } }}>
                {event.category}
              </Text>
            </Stack>
          )}
        </Stack>
      )}

      {/* ── Description ── */}
      {event.description && (
        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 16 } }}>
          <Text
            variant="tiny"
             styles={{
              root: { color: '#8a8886', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 },
            }}
          >
            Description
          </Text>
          <div
            style={{
              border: '1px solid #edebe9',
              borderRadius: 6,
              padding: '12px 16px',
              background: '#faf9f8',
              fontSize: 14,
              lineHeight: 1.6,
              color: '#323130',
              maxHeight: 200,
              overflowY: 'auto',
            }}
            dangerouslySetInnerHTML={{ __html: event.description }}
          />
        </Stack>
      )}

      {/* ── Recurrence ── */}
      {event.RecurrenceData && (
        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 16 } }}>
          <Text
            variant="tiny"
              styles={{
              root: { color: '#8a8886', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 },
            }}
          >
            Recurrence
          </Text>
          <Stack
            horizontal
            verticalAlign="center"
            tokens={{ childrenGap: 8 }}
            styles={{
              root: {
                background: '#eff6fc',
                border: '1px solid #c7e0f4',
                borderRadius: 6,
                padding: '10px 14px',
              },
            }}
          >
            <span style={{ fontSize: 16 }}>🔁</span>
            <Text variant="small" styles={{ root: { color: '#004e8c' } }}>
              {parseRecurrenceToString(event.RecurrenceData)}
            </Text>
          </Stack>
        </Stack>
      )}

      {/* ── PDF Attachments — inline iframe preview (matches extension style) ── */}
      {pdfAttachments.length > 0 && (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 8 } }}>
          <Text
            variant="tiny"
            styles={{
              root: { color: '#8a8886', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 },
            }}
          >
            Attached Documents
          </Text>
          {pdfAttachments.map(a => (
            <Stack
              key={a.FileName}
              tokens={{ childrenGap: 0 }}
              styles={{
                root: {
                  border: '1px solid #edebe9',
                  borderRadius: 6,
                  overflow: 'hidden',
                  boxShadow: '0 1px 4px rgba(0,0,0,0.07)',
                },
              }}
            >
              {/* Title bar — mirrors extension popup chrome */}
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: {
                    background: '#f3f2f1',
                    borderBottom: '1px solid #edebe9',
                    padding: '8px 12px',
                  },
                }}
              >
                <span style={{ fontSize: 16 }}>📄</span>
                <Text
                  styles={{
                    root: {
                      flex: 1,
                      fontWeight: 600,
                      fontSize: 13,
                      color: '#323130',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    },
                  }}
                >
                  {a.FileName}
                </Text>
                <IconButton
                  iconProps={{ iconName: 'OpenInNewWindow' }}
                  title="Open in new tab"
                  href={a.ServerRelativeUrl}
                  target="_blank"
                  styles={{
                    root: { height: 28, width: 28, color: '#0078d4' },
                    icon: { fontSize: 14 },
                  }}
                />
              </Stack>

              {/* Inline PDF iframe — same approach as extension */}
              <iframe
                src={`${a.ServerRelativeUrl}#toolbar=0&navpanes=0&scrollbar=1`}
                title={a.FileName}
                style={{
                  width: '100%',
                  height: 500,
                  border: 'none',
                  display: 'block',
                  background: '#525659',
                }}
              />
            </Stack>
          ))}
        </Stack>
      )}

      <DialogFooter styles={{ actions: { marginTop: 16 } }}>
        <DefaultButton
          text="Close"
          onClick={onClose}
          styles={{ root: { minWidth: 90, borderRadius: 4 } }}
        />
      </DialogFooter>
    </Dialog>
  );
};

// ─── Main Component ───────────────────────────────────────────────────────────
const EventForm: React.FC<any> = (props) => {
  const sp = props?.siteUrl != null
    ? spfi(props.siteUrl).using(SPFx(props.Context))
    : spfi().using(SPFx(props.Context));

  const canEdit = props.isNew ? !!props?.userPermissions?.canAdd : !!props?.canEditEvent;

  // ─── Form State ───────────────────────────────────────────────────────────────
  const [isAllDay, setIsAllDay] = useState(false);
  const [formData, setFormData] = useState<any>({
    id: '',
    title: '',
    locations: '',
    startTime: new Date(),
    endTime: new Date(Date.now() + 60 * 60 * 1000),
    description: '',
    attendees: [],
    category: '',
    fAllDayEvent: false,
    resources: '',
    freeBusy: 'Busy',
    checkDoubleBooking: false,
    modified: new Date(),
    created: new Date(),
    createdBy: null,
    modifiedBy: null,
    RecurrenceData: '',
  });

  // ─── Preview state ────────────────────────────────────────────────────────────
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);

  const [showRecurrenceSeriesInfo, setShowRecurrenceSeriesInfo] = useState(false);
  const [showRecurrence] = useState(true);
  const [IsDisableField] = useState(false);
  const [categoryOptions, setCategoryOptions] = useState([
    { key: 'Meeting', text: 'Meeting' },
    { key: 'RFQ', text: 'RFQ' },
    { key: 'RFP', text: 'RFP' },
    { key: 'CSP/Traditional', text: 'CSP/Traditional' },
    { key: 'DB', text: 'DB' },
    { key: 'Interview', text: 'Interview' },
  ]);
  const [returnedRecurrenceInfo, setReturnedRecurrenceInfo] =
    useState<{ recurrenceData: string; eventDate: Date; endDate: Date } | null>(null);
  const [recurrenceData, setRecurrenceData] = useState<string | null>(null);
  const [isEditingRecurrence, setIsEditingRecurrence] = useState(false);
  const [tempRecurrenceData, setTempRecurrenceData] = useState<any>(null);
  const [timeError, setTimeError] = useState<string | null>(null);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [saveError, setSaveError] = useState<string | null>(null);

  // ─── Attachment State ─────────────────────────────────────────────────────────
  const [existingAttachments, setExistingAttachments] = useState<IAttachmentInfo[]>([]);
  const [pendingLocalFile, setPendingLocalFile] = useState<File | null>(null);
  const [pendingLibraryBlob, setPendingLibraryBlob] = useState<{ blob: Blob; name: string } | null>(null);
  const [attachmentsToDelete, setAttachmentsToDelete] = useState<string[]>([]);

  // ─── PDF Picker State ─────────────────────────────────────────────────────────
  const [isPdfPickerOpen, setIsPdfPickerOpen] = useState(false);
  const [pdfCurrentFolder, setPdfCurrentFolder] = useState(PDF_FOLDER_SERVER_RELATIVE);
  const [pdfFolderTrail, setPdfFolderTrail] = useState<IBreadcrumbItem[]>([]);
  const [pdfSubfolders, setPdfSubfolders] = useState<any[]>([]);
  const [pdfFiles, setPdfFiles] = useState<any[]>([]);
  const [pdfPickerError, setPdfPickerError] = useState<string | null>(null);
  const [pdfUploading, setPdfUploading] = useState(false);
  const [pdfFetching, setPdfFetching] = useState(false);

  // ─── Image Picker State ───────────────────────────────────────────────────────
  const [isImagePickerOpen, setIsImagePickerOpen] = useState(false);
  const [imgCurrentFolder, setImgCurrentFolder] = useState(IMG_FOLDER_SERVER_RELATIVE);
  const [imgFolderTrail, setImgFolderTrail] = useState<IBreadcrumbItem[]>([]);
  const [imgSubfolders, setImgSubfolders] = useState<any[]>([]);
  const [imgFiles, setImgFiles] = useState<any[]>([]);
  const [imgPickerError, setImgPickerError] = useState<string | null>(null);
  const [imgUploading, setImgUploading] = useState(false);

  // ─── Effects ──────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!props.event) return;
    setFormData({ ...props.event });
    if (props.event.fAllDayEvent) setIsAllDay(true);
    if (props.event.RecurrenceData) {
      setShowRecurrenceSeriesInfo(true);
      setRecurrenceData(props.event.RecurrenceData);
    }
    if (!props.isNew && props.event.id) {
      loadExistingAttachments(props.event.id);
    }

    if (props?.CalendarTitle === 'Marketing Calendar-Internal') {
      setCategoryOptions([
        { key: 'PTO', text: 'PTO' },
        { key: 'Remote', text: 'Remote' },
        { key: 'Appointment', text: 'Appointment' },
        { key: 'Site Visit', text: 'Site Visit' },
        { key: 'Interview Prep', text: 'Interview Prep' },
        { key: 'Other', text: 'Other' },
      ]);
    } else if (props?.CalendarTitle === 'Company Calendar') {
      setCategoryOptions([
        { key: 'Safety', text: 'Safety' },
        { key: 'Ops', text: 'Ops' },
        { key: 'Other', text: 'Other' },
        { key: 'Vaughn Outdoors', text: 'Vaughn Outdoors' },
        { key: 'VaughnLife', text: 'VaughnLife' },
        { key: 'HR', text: 'HR' },
      ]);
    }
  }, [props.event]);

  useEffect(() => { validateDates(); }, [formData.startTime, formData.endTime]);

  useEffect(() => {
    if (isPdfPickerOpen) loadPdfFolder(pdfCurrentFolder);
  }, [isPdfPickerOpen, pdfCurrentFolder]);

  useEffect(() => {
    if (isImagePickerOpen) loadImageFolder(imgCurrentFolder);
  }, [isImagePickerOpen, imgCurrentFolder]);

  // ─── Loaders ──────────────────────────────────────────────────────────────────
  const loadExistingAttachments = async (itemId: string | number) => {
    try {
      const attachments: IAttachmentInfo[] = await sp.web
        .lists.getById(props.MarketingCalendarId)
        .items.getById(parseInt(String(itemId)))
        .attachmentFiles();
      setExistingAttachments(attachments);
    } catch (e) {
      console.error('Error loading attachments:', e);
    }
  };

  const loadPdfFolder = async (folderPath: string) => {
    try {
      setPdfPickerError(null);
      const spPdf = spfi(PDF_SP_ROOT).using(SPFx(props.Context));
      const folder = spPdf.web.getFolderByServerRelativePath(PDF_FOLDER_SERVER_RELATIVE);
      const [subs, files] = await Promise.all([folder.folders(), folder.files()]);
      setPdfSubfolders(subs);
      setPdfFiles(files);
      setPdfFolderTrail(buildBreadcrumb(folderPath, setPdfCurrentFolder));
    } catch (e) {
      setPdfPickerError(parseError(e));
    }
  };

  const loadImageFolder = async (folderPath: string) => {
    try {
      setImgPickerError(null);
      const spImg = spfi(IMG_SP_ROOT).using(SPFx(props.Context));
      const folder = spImg.web.getFolderByServerRelativePath(IMG_FOLDER_SERVER_RELATIVE);
      const [subs, files] = await Promise.all([folder.folders(), folder.files()]);
      setImgSubfolders(subs);
      setImgFiles(files);
      setImgFolderTrail(buildBreadcrumb(folderPath, setImgCurrentFolder));
    } catch (e) {
      setImgPickerError(parseError(e));
    }
  };

  // ─── Validation ───────────────────────────────────────────────────────────────
  const validateDates = () => {
    if (formData.endTime < formData.startTime) {
      const newEndTime = new Date(formData.startTime.getTime() + 60 * 60 * 1000);
      setFormData((prev: any) => ({ ...prev, endTime: newEndTime }));
      setTimeError('End time must be after start time. It has been automatically adjusted.');
    } else {
      setTimeError(null);
    }
  };

  const handleInputChange = (field: string, value: any): void => {
    setFormData((prev: any) => ({ ...prev, [field]: value }));
    if (errors[field]) setErrors((prev: any) => ({ ...prev, [field]: '' }));
  };

  const handleRichTextChange = (text: string): string => {
    handleInputChange('description', text);
    return text;
  };

  const handleTimeChange = (field: 'startTime' | 'endTime', newTime: Date | null | undefined) => {
    if (!newTime) return;
    try {
      const currentDate: Date = formData[field];
      const updatedDate = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth(),
        currentDate.getDate(),
        newTime.getHours(),
        newTime.getMinutes(),
        0
      );
      handleInputChange(field, updatedDate);
      if (field === 'startTime' && updatedDate >= formData.endTime) {
        handleInputChange('endTime', new Date(updatedDate.getTime() + 60 * 60 * 1000));
      }
    } catch (e) {
      console.error(`Error updating ${field}:`, e);
    }
  };

  // ─── PDF Handlers ─────────────────────────────────────────────────────────────
  const handlePdfUploadToLibrary = async (files: FileList | null) => {
    if (!files?.length) return;
    try {
      setPdfUploading(true);
      setPdfPickerError(null);
      const spPdf = spfi(PDF_SP_ROOT).using(SPFx(props.Context));
      const file = files[0];
      const folder = spPdf.web.getFolderByServerRelativePath(PDF_FOLDER_SERVER_RELATIVE);
      await folder.files.addUsingPath(file.name.trim(), file, { Overwrite: true });
      await loadPdfFolder(pdfCurrentFolder);
    } catch (e) {
      setPdfPickerError(parseError(e));
    } finally {
      setPdfUploading(false);
    }
  };

  const handleSelectPdfFromLibrary = async (serverRelativeUrl: string, fileName: string) => {
    try {
      setPdfFetching(true);
      setPdfPickerError(null);
      const spPdf = spfi(PDF_SP_ROOT).using(SPFx(props.Context));
      const blob: Blob = await spPdf.web.getFileByServerRelativePath(serverRelativeUrl).getBlob();
      setPendingLibraryBlob({ blob, name: fileName });
      setPendingLocalFile(null);
      setIsPdfPickerOpen(false);
    } catch (e) {
      setPdfPickerError(parseError(e));
    } finally {
      setPdfFetching(false);
    }
  };

  const handleLocalPdfFileChange = (files: FileList | null) => {
    if (!files?.length) return;
    setPendingLocalFile(files[0]);
    setPendingLibraryBlob(null);
  };

  const markAttachmentForDeletion = (fileName: string) => {
    setAttachmentsToDelete(prev => [...prev, fileName]);
    setExistingAttachments(prev => prev.filter(a => a.FileName !== fileName));
  };

  // ─── Image Handlers ───────────────────────────────────────────────────────────
  const handleImageUpload = async (files: FileList | null) => {
    if (!files?.length) return;
    try {
      setImgUploading(true);
      setImgPickerError(null);
      const spImg = spfi(IMG_SP_ROOT).using(SPFx(props.Context));
      const file = files[0];
      const folder = spImg.web.getFolderByServerRelativePath(IMG_FOLDER_SERVER_RELATIVE);
      await folder.files.addUsingPath(file.name.trim(), file, { Overwrite: true });
      await loadImageFolder(imgCurrentFolder);
    } catch (e) {
      setImgPickerError(parseError(e));
    } finally {
      setImgUploading(false);
    }
  };

  const insertImageIntoDescription = (serverRelativeUrl: string, name: string) => {
    const imgTag = `<img src="${serverRelativeUrl}" alt="${name}" style="max-width:100%;height:auto;" />`;
    setFormData((prev: any) => ({ ...prev, description: (prev.description || '') + imgTag }));
    setIsImagePickerOpen(false);
  };

  // ─── Save ─────────────────────────────────────────────────────────────────────
 const handleSave = async () => {
  if (props.isNew && !props?.userPermissions?.canAdd) return;
  if (!props.isNew && !props?.canEditEvent) return;

  if (formData.endTime < formData.startTime) {
    setTimeError('End time must be after start time.');
    return;
  }

  setSaveError(null);

  try {
    const eventToSave = { ...formData } as any;

    if (returnedRecurrenceInfo) {
      eventToSave.RecurrenceData = returnedRecurrenceInfo.recurrenceData;
    } else if (recurrenceData) {
      eventToSave.RecurrenceData = recurrenceData;
    }

    const listItem: any = {
      Title: eventToSave.title,
      Location: eventToSave.locations,
      Description: eventToSave.description,
      Category: eventToSave.category,
      fAllDayEvent: isAllDay,
    };

    if (isAllDay) {
      const dateOnly = moment(eventToSave.startTime).format('YYYY-MM-DD');
      listItem.EventDate = `${dateOnly}T00:00:00`;
      listItem.EndDate = `${dateOnly}T23:59:59`;
      listItem.fAllDayEvent = true;
    } else {
      listItem.EventDate = moment(eventToSave.startTime).format('YYYY-MM-DDTHH:mm:ss') + 'Z';
      listItem.EndDate = moment(eventToSave.endTime).format('YYYY-MM-DDTHH:mm:ss') + 'Z';
    }

    if (eventToSave.RecurrenceData) {
      listItem.RecurrenceData = eventToSave.RecurrenceData;
    }

    let savedItemId: number;

    // ─────────────────────────────────────────
    // CREATE OR UPDATE EVENT
    // ─────────────────────────────────────────
    if (props.isNew) {

      const response = await sp.web.lists
        .getById(props.MarketingCalendarId)
        .items.add(listItem);

      savedItemId = response?.Id;
      console.log('Item added successfully:', savedItemId);

      // 🔴 IMPORTANT: wait for SharePoint commit
      await new Promise(resolve => setTimeout(resolve, 1200));

    } else {

      savedItemId = parseInt(String(eventToSave.id));

      await sp.web.lists
        .getById(props.MarketingCalendarId)
        .items.getById(savedItemId)
        .update(listItem);

      console.log('Item updated successfully');

    }

    // ─────────────────────────────────────────
    // ATTACHMENTS
    // ─────────────────────────────────────────
    const itemAttachments = sp.web.lists
      .getById(props.MarketingCalendarId)
      .items
      .getById(savedItemId)
      .attachmentFiles;

    // Delete attachments
    for (const fileName of attachmentsToDelete) {
      try {
        await itemAttachments.getByName(fileName).delete();
      } catch (e) {
        console.warn(`Could not delete attachment "${fileName}":`, e);
      }
    }

    // Upload local file
    if (pendingLocalFile) {

      const buffer = await pendingLocalFile.arrayBuffer();

      await itemAttachments.add(
        pendingLocalFile.name,
        buffer
      );

      setPendingLocalFile(null);
    }

    // Upload library selected file
    if (pendingLibraryBlob) {

      const buffer = await pendingLibraryBlob.blob.arrayBuffer();

      await itemAttachments.add(
        pendingLibraryBlob.name,
        buffer
      );

      setPendingLibraryBlob(null);
    }

    // Reload attachments
    const updatedAttachments: IAttachmentInfo[] = await itemAttachments();

    setExistingAttachments(updatedAttachments);
    setAttachmentsToDelete([]);

    props.onSave({
      ...eventToSave,
      id: savedItemId,
      attachments: updatedAttachments
    });

  } catch (error) {

    console.error('Error saving event:', error);

    setSaveError(parseError(error));
  }
};

  const handleDelete = (): void => {
    if (formData.id && !props.isNew) props.onDelete(formData.id);
  };

  // ─── Recurrence Helpers ───────────────────────────────────────────────────────
  const returnRecurrenceInfo = (startDate: Date, endDate: Date, recurrenceDataVal: string) => {
    const info = { recurrenceData: recurrenceDataVal, eventDate: startDate, endDate };
    if (props.isNew || !props.event?.RecurrenceData) {
      setFormData((prev: any) => ({ ...prev, RecurrenceData: info.recurrenceData }));
      setReturnedRecurrenceInfo(info);
    } else {
      setTempRecurrenceData(info);
    }
  };

  const handleRecurrenceCheck = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, recurChecked: boolean) => {
    ev.preventDefault();
    setShowRecurrenceSeriesInfo(recurChecked);
  };

  const getRecurrenceType = (rd: any): string => {
    if (!rd) return '';
    if (rd.indexOf('<daily') !== -1) return 'daily';
    if (rd.indexOf('<weekly') !== -1) return 'weekly';
    if (rd.indexOf('<monthly') !== -1) return 'monthly';
    if (rd.indexOf('<yearly') !== -1) return 'yearly';
    return '';
  };

  const saveRecurrenceChanges = () => {
    setRecurrenceData(tempRecurrenceData?.recurrenceData);
    setReturnedRecurrenceInfo(tempRecurrenceData);
    setFormData((prev: any) => ({ ...prev, RecurrenceData: tempRecurrenceData?.recurrenceData }));
    setIsEditingRecurrence(false);
  };

  const cancelRecurrenceChanges = () => {
    setTempRecurrenceData(returnedRecurrenceInfo);
    setIsEditingRecurrence(false);
  };

  const startEditingRecurrence = () => {
    setTempRecurrenceData(recurrenceData);
    setIsEditingRecurrence(true);
  };

  // ─── Pending label ────────────────────────────────────────────────────────────
  const pendingLabel = pendingLocalFile
    ? pendingLocalFile.name
    : pendingLibraryBlob
    ? pendingLibraryBlob.name
    : null;

  // ─── Description field ────────────────────────────────────────────────────────
  const renderDescriptionField = () => {
    if (props?.CalendarTitle === 'Company Calendar') {
      return (
        <>
          <RichText value={formData.description} onChange={handleRichTextChange} isEditMode={true} />
          {errors.description && (
            <Text variant="small" style={{ color: '#d13438' }}>{errors.description}</Text>
          )}
        </>
      );
    }
    return (
      <TextField
        multiline
        rows={3}
        value={formData.description}
        onChange={(_, value) => handleInputChange('description', value)}
        styles={{ fieldGroup: { borderRadius: 4 } }}
      />
    );
  };

  // ─── Footer: Preview (left) + Save / Delete / Cancel (right) ─────────────────

  // ADD THIS — resolves bg/fg for any category, mirrors ModernCalendar's maps exactly
const categoryOptionsColor: Record<string, string> = {
  'Meeting': '#3174ad', 'RFQ': '#ffff00', 'RFP': '#107c10',
  'CSP/Traditional': '#da3b01', 'DB': '#c239b3', 'Interview': '#adadad',
  'Appointment': '#0099bc', 'Site Visit': '#00b294', 'Remote': '#004e8c',
  'Interview Prep': '#ffaa44', 'Other': '#605e5c', 'PTO': '#e3008c',
  'N/A': '#605e5c',
};
const categoryOptionsFontColor: Record<string, string> = {
  'Meeting': '#ffffff', 'RFQ': '#000000', 'RFP': '#ffffff',
  'CSP/Traditional': '#ffffff', 'DB': '#ffffff', 'Interview': '#000000',
  'Appointment': '#ffffff', 'Site Visit': '#000000', 'Remote': '#ffffff',
  'Interview Prep': '#000000', 'Other': '#ffffff', 'PTO': '#ffffff',
  'N/A': '#ffffff',
};
const categoryOptionsColorTraining: Record<string, string> = {
  'Safety': '#ff0000', 'Ops': '#3174ad', 'Other': '#ffff00',
  'Vaughn Outdoors': '#8b4513', 'VaughnLife': '#107c10', 'HR': '#ff8c00',
  'N/A': '#605e5c',
};
const categoryOptionsFontColorTraining: Record<string, string> = {
  'Safety': '#ffffff', 'Ops': '#ffffff', 'Other': '#000000',
  'Vaughn Outdoors': '#ffffff', 'VaughnLife': '#ffffff', 'HR': '#ffffff',
  'N/A': '#ffffff',
};
const isTrainingCalendar = props?.CalendarTitle === 'Company Calendar';
const resolvedCategoryBg = isTrainingCalendar
  ? (categoryOptionsColorTraining[formData.category] ?? '#605e5c')
  : (categoryOptionsColor[formData.category] ?? '#605e5c');
const resolvedCategoryFg = isTrainingCalendar
  ? (categoryOptionsFontColorTraining[formData.category] ?? '#ffffff')
  : (categoryOptionsFontColor[formData.category] ?? '#ffffff');
  const renderFooterContent = () => (
    <Stack
      horizontal
      tokens={{ childrenGap: 8 }}
      horizontalAlign="space-between"
      verticalAlign="center"
      styles={{ root: { width: '100%' } }}
    >
      <DefaultButton
        text="Preview"
        iconProps={{ iconName: 'View' }}
        onClick={() => setIsPreviewOpen(true)}
        styles={{
          root: { borderColor: '#0078d4', color: '#0078d4', borderRadius: 4 },
          icon: { color: '#0078d4' },
        }}
      />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton onClick={handleSave} text="Save" styles={{ root: { borderRadius: 4 } }} />
        {!props.isNew && props?.canDeleteEvent && (
          <DefaultButton
            onClick={handleDelete}
            text="Delete"
            styles={{ root: { borderRadius: 4, color: '#d13438', borderColor: '#d13438' } }}
          />
        )}
        <DefaultButton onClick={props.onCancel} text="Cancel" styles={{ root: { borderRadius: 4 } }} />
      </Stack>
    </Stack>
  );

  // ─── Read-only view ───────────────────────────────────────────────────────────
  if (!canEdit && !props.isNew) {
    const viewEvent = formData.id ? formData : props.event;
    if (!viewEvent) return null;
     return (
      <EventViewDialog
        event={viewEvent}
        attachments={existingAttachments}
        onClose={props.onCancel}
        categoryBg={resolvedCategoryBg}
        categoryFg={resolvedCategoryFg}
      />
    );
  }

  // ─── Edit Panel ───────────────────────────────────────────────────────────────
  return (
    <>
      <Panel
        isBlocking={false}
        isOpen={true}
        onDismiss={props.onCancel}
        headerText={props.isNew ? 'Add Event' : 'Edit Event'}
        type={PanelType.medium}
        onRenderFooterContent={renderFooterContent}
        isFooterAtBottom={true}
      >
        <Stack tokens={stackTokens} styles={stackStyles}>

          {saveError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setSaveError(null)}>
              {saveError}
            </MessageBar>
          )}

          <TextField
            label="Title"
            value={formData.title}
            onChange={(_, value) => handleInputChange('title', value)}
            required
            styles={{ fieldGroup: { borderRadius: 4 } }}
          />

          <TextField
            label="Location"
            value={formData.locations}
            onChange={(_, value) => handleInputChange('locations', value)}
            styles={{ fieldGroup: { borderRadius: 4 } }}
          />

          {timeError && (
            <MessageBar
              messageBarType={MessageBarType.warning}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
              onDismiss={() => setTimeError(null)}
            >
              {timeError}
            </MessageBar>
          )}

          <Toggle
            label="All Day Event"
            checked={isAllDay}
            onChange={(_, checked) => {
              setIsAllDay(!!checked);
              if (checked) {
                const start = new Date(formData.startTime);
                start.setHours(0, 0, 0, 0);
                const end = new Date(formData.endTime);
                end.setHours(23, 59, 0, 0);
                handleInputChange('startTime', start);
                handleInputChange('endTime', end);
              }
            }}
          />

          {/* ── Start Date / Time ── */}
          <Stack horizontal tokens={stackTokens}>
            <Stack.Item grow={1}>
              <DatePicker
                label="Start Date"
                value={formData.startTime}
                onSelectDate={(date) => {
                  if (date) {
                    const newDate = new Date(formData.startTime);
                    newDate.setFullYear(date.getFullYear(), date.getMonth(), date.getDate());
                    handleInputChange('startTime', newDate);
                  }
                }}
                firstDayOfWeek={DayOfWeek.Sunday}
              />
            </Stack.Item>
            {!isAllDay && (
              <Stack.Item grow={1}>
                <TimePicker
                  label="Start Time"
                  value={formData.startTime}
                  onChange={(_ev: any, time: Date | null | undefined) => handleTimeChange('startTime', time)}
                  increments={15}
                />
              </Stack.Item>
            )}
          </Stack>

          {/* ── End Date / Time ── */}
          <Stack horizontal tokens={stackTokens}>
            <Stack.Item grow={1}>
              <DatePicker
                label="End Date"
                value={formData.endTime}
                onSelectDate={(date) => {
                  if (date) {
                    const startDay = new Date(formData.startTime);
                    startDay.setHours(0, 0, 0, 0);
                    if (date >= startDay) {
                      const newDate = new Date(formData.endTime);
                      newDate.setFullYear(date.getFullYear(), date.getMonth(), date.getDate());
                      handleInputChange('endTime', newDate);
                    }
                  }
                }}
                firstDayOfWeek={DayOfWeek.Sunday}
                minDate={formData.startTime}
              />
            </Stack.Item>
            {!isAllDay && (
              <Stack.Item grow={1}>
                <TimePicker
                  label="End Time"
                  value={formData.endTime}
                  onChange={(_ev: any, time: Date | null | undefined) => handleTimeChange('endTime', time)}
                  increments={15}
                />
              </Stack.Item>
            )}
          </Stack>

          {/* ── Description ── */}
          <Stack tokens={{ childrenGap: 6 }}>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <Label styles={{ root: { marginBottom: 0, flex: 1 } }}>Description</Label>
              <DefaultButton
                text="Insert Image"
                iconProps={{ iconName: 'Photo' }}
                onClick={() => setIsImagePickerOpen(true)}
                styles={{ root: { height: 28, fontSize: 13, borderRadius: 4 } }}
              />
            </Stack>
            {renderDescriptionField()}
          </Stack>

          <Dropdown
            label="Category"
            options={categoryOptions}
            selectedKey={formData.category}
            onChange={(_, option) => option && handleInputChange('category', option.key)}
          />

          {/* ── PDF Attachment ── */}
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>PDF Attachment</Label>

            {existingAttachments.filter(a => a.FileName.toLowerCase().endsWith('.pdf')).length > 0 && (
              <Stack tokens={{ childrenGap: 6 }}>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>Saved attachments:</Text>
                {existingAttachments
                  .filter(a => a.FileName.toLowerCase().endsWith('.pdf'))
                  .map(a => (
                    <Stack
                      key={a.FileName}
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: {
                          padding: '6px 10px',
                          border: '1px solid #edebe9',
                          borderRadius: 4,
                          background: '#faf9f8',
                        },
                      }}
                    >
                      <span>📄</span>
                      <Text styles={{ root: { flex: 1 } }}>{a.FileName}</Text>
                      <IconButton
                        iconProps={{ iconName: 'OpenInNewWindow' }}
                        title="Open PDF"
                        href={a.ServerRelativeUrl}
                        target="_blank"
                      />
                      <DefaultButton
                        text="Remove"
                        styles={{ root: { height: 28, borderRadius: 4 } }}
                        onClick={() => markAttachmentForDeletion(a.FileName)}
                      />
                    </Stack>
                  ))}
              </Stack>
            )}

            {pendingLabel && (
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: { padding: '6px 10px', border: '1px dashed #0078d4', borderRadius: 4 },
                }}
              >
                <span>📎</span>
                <Text styles={{ root: { flex: 1 } }}>
                  {pendingLabel}{' '}
                  <em style={{ color: '#605e5c', fontSize: 12 }}>(pending — will be attached on Save)</em>
                </Text>
                <DefaultButton
                  text="Clear"
                  styles={{ root: { height: 28, borderRadius: 4 } }}
                  onClick={() => { setPendingLocalFile(null); setPendingLibraryBlob(null); }}
                />
              </Stack>
            )}

            <Stack horizontal tokens={{ childrenGap: 8 }} wrap verticalAlign="center">
                <Stack.Item grow={1}>
<DefaultButton
                text="Select from Memos Library"
                iconProps={{ iconName: 'PDF' }}
                onClick={() => setIsPdfPickerOpen(true)}
                styles={{ root: { borderRadius: 4 } }}
              />
                </Stack.Item>
                  <Stack.Item grow={1}>
<Label styles={{ root: { margin: 0, whiteSpace: 'nowrap' } }}>Upload from computer:</Label>
                <input
                  type="file"
                  accept="application/pdf"
                  onChange={(e) => handleLocalPdfFileChange(e.target.files)}
                  style={{ fontSize: 13 }}
                />
                  </Stack.Item>
              
              {/* <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                
              </Stack> */}
            </Stack>
          </Stack>

          {/* ── Recurrence ── */}
          {showRecurrence && (
            <Toggle
              label="Recurrence"
              checked={showRecurrenceSeriesInfo}
              onChange={handleRecurrenceCheck}
              disabled={IsDisableField}
            />
          )}

          {showRecurrenceSeriesInfo && !props.isNew && recurrenceData && !isEditingRecurrence && (
            <Stack tokens={stackTokens}>
              <Label>Recurrence Pattern</Label>
              <Stack horizontal verticalAlign="center">
                <Text>{parseRecurrenceToString(recurrenceData)}</Text>
                <IconButton
                  iconProps={{ iconName: 'Edit' }}
                  title="Edit recurrence"
                  ariaLabel="Edit recurrence"
                  onClick={startEditingRecurrence}
                  styles={{ root: { marginLeft: 8 } }}
                />
              </Stack>
            </Stack>
          )}

          {showRecurrenceSeriesInfo && (isEditingRecurrence || props.isNew || !recurrenceData) && (
            <Stack>
              <EventRecurrenceInfo
                context={props.Context}
                startDate={formData.startTime}
                endDate={formData.endTime}
                returnRecurrenceInfo={returnRecurrenceInfo}
                siteUrl={props?.siteUrl ?? props.Context.pageContext.web.absoluteUrl}
                recurrenceData={!isEditingRecurrence ? recurrenceData : ''}
                removeRecurrence={true}
                selectedRecurrenceRule={
                  !isEditingRecurrence ? getRecurrenceType(returnedRecurrenceInfo?.recurrenceData) : ''
                }
                selectedKey={
                  !isEditingRecurrence ? getRecurrenceType(returnedRecurrenceInfo?.recurrenceData) : ''
                }
                display={true}
                DueDate={formData.endTime}
              />
              {isEditingRecurrence && !props.isNew && (
                <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                  <PrimaryButton text="Update Recurrence" onClick={saveRecurrenceChanges} />
                  <DefaultButton text="Cancel" onClick={cancelRecurrenceChanges} />
                </Stack>
              )}
            </Stack>
          )}
        </Stack>
      </Panel>

      {/* ─── Preview Dialog — opened from edit form ──────────────────────────── */}
       {isPreviewOpen && (
        <EventViewDialog
          event={{ ...formData, fAllDayEvent: isAllDay,
            RecurrenceData: returnedRecurrenceInfo?.recurrenceData || recurrenceData || formData.RecurrenceData }}
          attachments={existingAttachments}
          onClose={() => setIsPreviewOpen(false)}
          categoryBg={resolvedCategoryBg}
          categoryFg={resolvedCategoryFg}
        />
      )}
      {/* ─── PDF Library Picker Panel ─────────────────────────────────────────── */}
      <Panel
        isOpen={isPdfPickerOpen}
        onDismiss={() => setIsPdfPickerOpen(false)}
        type={PanelType.large}
        isBlocking={true}
        headerText="Select PDF — Memos Library"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          {pdfPickerError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setPdfPickerError(null)}>
              {pdfPickerError}
            </MessageBar>
          )}

          {pdfFetching && <Spinner size={SpinnerSize.medium} label="Fetching PDF…" />}

          <Breadcrumb items={pdfFolderTrail} maxDisplayedItems={8} ariaLabel="PDF folder breadcrumb" overflowAriaLabel="More" />

          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <label style={{ fontWeight: 500 }}>Upload PDF to library:</label>
            <input type="file" accept="application/pdf" onChange={(e) => handlePdfUploadToLibrary(e.target.files)} disabled={pdfUploading} />
            {pdfUploading && <Spinner size={SpinnerSize.small} label="Uploading…" />}
          </Stack>

          <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 16 }}>
            <div>
              <h4 style={{ marginTop: 0 }}>Folders</h4>
              <ul style={{ paddingLeft: 16, listStyle: 'none', margin: 0 }}>
                {pdfSubfolders.map((f: any) => (
                  <li key={f.Name} style={{ marginBottom: 6 }}>
                    <a href="#" onClick={(e) => { e.preventDefault(); setPdfCurrentFolder(f.ServerRelativeUrl); }} style={{ textDecoration: 'none', color: '#0078d4' }}>
                      📁 {f.Name}
                    </a>
                  </li>
                ))}
                {pdfSubfolders.length === 0 && <li style={{ color: '#888' }}>No subfolders</li>}
              </ul>
            </div>

            <div>
              <h4 style={{ marginTop: 0 }}>PDF Files</h4>
              <Stack tokens={{ childrenGap: 8 }}>
                {pdfFiles
                  .filter((f: any) => f.Name?.toLowerCase().endsWith('.pdf'))
                  .map((file: any) => (
                    <Stack
                      key={file.UniqueId}
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 12 }}
                      styles={{ root: { padding: '8px 12px', border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8' } }}
                    >
                      <span style={{ fontSize: 22 }}>📄</span>
                      <Text styles={{ root: { flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }}>
                        {file.Name}
                      </Text>
                      <IconButton iconProps={{ iconName: 'OpenInNewWindow' }} title="Preview in new tab" href={file.ServerRelativeUrl} target="_blank" />
                      <PrimaryButton text="Select" disabled={pdfFetching} onClick={() => handleSelectPdfFromLibrary(file.ServerRelativeUrl, file.Name)} styles={{ root: { borderRadius: 4 } }} />
                    </Stack>
                  ))}
                {pdfFiles.filter((f: any) => f.Name?.toLowerCase().endsWith('.pdf')).length === 0 && (
                  <Text style={{ color: '#888' }}>No PDF files in this folder</Text>
                )}
              </Stack>
            </div>
          </div>
        </Stack>
      </Panel>

      {/* ─── Image Picker Panel ────────────────────────────────────────────────── */}
      <Panel
        isOpen={isImagePickerOpen}
        onDismiss={() => setIsImagePickerOpen(false)}
        type={PanelType.large}
        isBlocking={true}
        headerText="Select or Upload Image"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          {imgPickerError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setImgPickerError(null)}>
              {imgPickerError}
            </MessageBar>
          )}

          <Breadcrumb items={imgFolderTrail} maxDisplayedItems={8} ariaLabel="Image folder breadcrumb" overflowAriaLabel="More" />

          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <label style={{ fontWeight: 500 }}>Upload Image:</label>
            <input type="file" accept="image/*" onChange={(e) => handleImageUpload(e.target.files)} disabled={imgUploading} />
            {imgUploading && <Spinner size={SpinnerSize.small} label="Uploading…" />}
          </Stack>

          <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 16 }}>
            <div>
              <h4 style={{ marginTop: 0 }}>Folders</h4>
              <ul style={{ paddingLeft: 16, listStyle: 'none', margin: 0 }}>
                {imgSubfolders.map((f: any) => (
                  <li key={f.Name} style={{ marginBottom: 6 }}>
                    <a href="#" onClick={(e) => { e.preventDefault(); setImgCurrentFolder(f.ServerRelativeUrl); }} style={{ textDecoration: 'none', color: '#0078d4' }}>
                      📁 {f.Name}
                    </a>
                  </li>
                ))}
                {imgSubfolders.length === 0 && <li style={{ color: '#888' }}>No subfolders</li>}
              </ul>
            </div>

            <div>
              <h4 style={{ marginTop: 0 }}>Images</h4>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(140px, 1fr))', gap: 12 }}>
                {imgFiles
                  .filter((file: any) => IMAGE_EXTENSIONS.has(getFileExt(file.Name || '')))
                  .map((file: any) => (
                    <div key={file.UniqueId} style={{ border: '1px solid #eee', borderRadius: 8, padding: 8 }}>
                      <div style={{ aspectRatio: '1/1', overflow: 'hidden', borderRadius: 6, marginBottom: 6 }}>
                        <img src={file.ServerRelativeUrl} alt={file.Name} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 4 }}>
                        <span title={file.Name} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 80 }}>
                          {file.Name}
                        </span>
                        <DefaultButton
                          text="Use"
                          styles={{ root: { height: 28, minWidth: 40, padding: '0 8px', borderRadius: 4 } }}
                          onClick={() => insertImageIntoDescription(file.ServerRelativeUrl, file.Name)}
                        />
                      </div>
                    </div>
                  ))}
                {imgFiles.filter((f: any) => IMAGE_EXTENSIONS.has(getFileExt(f.Name || ''))).length === 0 && (
                  <Text style={{ color: '#888' }}>No images in this folder</Text>
                )}
              </div>
            </div>
          </div>
        </Stack>
      </Panel>
    </>
  );
};

export default EventForm;