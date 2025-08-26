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
  Text
} from '@fluentui/react';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { TimePicker } from '@fluentui/react/lib/TimePicker';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { ICalendarEvent } from './CalendarEvent';
import { EventRecurrenceInfo } from '../../../globalCommon/EventRecurrenceControls/EventRecurrenceInfo/EventRecurrenceInfo';
import { spfi, SPFx } from '@pnp/sp';
import { parseRecurrenceToString } from '../../../globalCommon/reccurenceStringToText';
import moment from 'moment';

const stackTokens: IStackTokens = { childrenGap: 12 };
const stackStyles: IStackStyles = {
  root: {
    padding: '0 0 12px 0',
  },
};

interface IEventFormProps {
  event: ICalendarEvent | null;
  isNew: boolean;
  Context: any;
  CalendarTitle?: any | null;
  userPermissions?: any | null;
  MarketingCalendarId: any;
  onSave: (event: ICalendarEvent) => void;
  onDelete: (eventId: string) => void;
  onCancel: () => void;
  canEditEvent?: any | null;
  canDeleteEvent?: any | null;
  siteUrl?: string;
}

const EventForm: React.FC<IEventFormProps> = (props) => {
  const sp = props?.siteUrl != undefined ? spfi(props?.siteUrl).using(SPFx(props?.Context)) : spfi().using(SPFx(props?.Context));
  const [isAllDay, setIsAllDay] = useState(false);
  const [formData, setFormData] = useState<ICalendarEvent>({
    id: '',
    title: '',
    locations: '',
    startTime: new Date(),
    endTime: new Date(new Date().getTime() + 60 * 60 * 1000), // Default end time is 1 hour after start
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
  const [showRecurrenceSeriesInfo, setShowRecurrenceSeriesInfo] = React.useState(false);
  const [showRecurrence, setshowRecurrence] = React.useState(true);
  const [IsDisableField, setIsDisableField] = React.useState(false);
  const [categoryOptions, setCategoryOptions] = React.useState([
    { key: 'Meeting', text: 'Meeting' },
    { key: 'RFQ', text: 'RFQ' },
    { key: 'RFP', text: 'RFP' },
    { key: 'CSP/Traditional', text: 'CSP/Traditional' },
    { key: 'DB', text: 'DB' },
    { key: 'Interview', text: 'Interview' }
  ]);
  const [newRecurrenceEvent, setNewRecurrenceEvent] = React.useState(false);
  const [editRecurrenceEvent, setEditRecurrenceEvent] = React.useState(false);
  const [returnedRecurrenceInfo, setReturnedRecurrenceInfo] =
    React.useState<{ recurrenceData: string; eventDate: Date; endDate: Date } | null>(null);
  const [recurrenceData, setRecurrenceData] = React.useState(null);
  const [isEditingRecurrence, setIsEditingRecurrence] = useState<boolean>(false);
  const [tempRecurrenceData, setTempRecurrenceData] = useState<any | undefined>();
  const [timeError, setTimeError] = useState<string | null>(null);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});

  useEffect(() => {
    if (props.event) {
      setFormData({ ...props.event });
      if (props.event.fAllDayEvent === true) {
        setIsAllDay(true);
      }
      // Check if this is a recurring event by looking for RecurrenceData
      if (props.event.RecurrenceData) {
        setShowRecurrenceSeriesInfo(true);
        setRecurrenceData(props.event.RecurrenceData);
        setNewRecurrenceEvent(false);  // Not a new recurrence, editing existing
        setEditRecurrenceEvent(true);  // Flag as editing recurrence
      }
    }
    if (props?.CalendarTitle == 'Marketing Calendar-Internal') {
      setCategoryOptions([
        { key: 'PTO', text: 'PTO' },
        { key: 'Remote', text: 'Remote' },
        { key: 'Appointment', text: 'Appointment' },
        { key: 'Site Visit', text: 'Site Visit' },
        { key: 'Interview Prep', text: 'Interview Prep' },
        { key: 'Other', text: 'Other' }
      ]);
    } else if (props?.CalendarTitle == 'Company Calendar') {
      setCategoryOptions([
        { key: 'Safety', text: 'Safety' },
        { key: 'Ops', text: 'Ops' },
        { key: 'HR', text: 'HR' },
        { key: 'Other', text: 'Other' }
      ]);
    }
  }, [props.event]);

  // Validate dates whenever they change
  useEffect(() => {
    validateDates();
  }, [formData.startTime, formData.endTime]);

  const validateDates = () => {
    if (formData.endTime < formData.startTime) {
      // Automatically fix the end date/time to be after start date/time
      const newEndTime = new Date(formData.startTime.getTime() + 60 * 60 * 1000); // 1 hour after start time
      setFormData(prevData => ({
        ...prevData,
        endTime: newEndTime
      }));
      setTimeError("End time must be after start time. It has been automatically adjusted.");
    } else {
      setTimeError(null);
    }
  };

  const handleInputChange = (field: keyof ICalendarEvent, value: any): void => {
    setFormData(prevData => ({
      ...prevData,
      [field]: value
    }));

    // Clear error for this field when user starts typing
    if (errors[field]) {
      setErrors(prev => ({
        ...prev,
        [field]: ''
      }));
    }
  };

  const handleRichTextChange = (text: string): string => {
    handleInputChange('description', text);
    return text;
  };

  const getUtcTime = async (date: string | Date): Promise<string> => {
    try {
      // Initialize PnPjs SPFI with your site URL and properly set up the observer

      // Ensure date is a Date object
      const dateObj = typeof date === "string" ? new Date(date) : date;

      // Convert local time to UTC using the web's regional settings
      const utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(dateObj);

      // Return as ISO string, or format as needed
      return new Date(utcTime).toISOString();
    } catch (error) {
      console.error("Error converting time to UTC:", error);
      return Promise.reject(error);
    }
  }

  const handleSave = async () => {
    // Validate dates before saving
    if (props?.isNew && !props?.userPermissions.canAdd) {
      console.warn('User does not have permission to add events');
      return;
    }

    if (!props?.isNew && !props?.canEditEvent) {
      console.warn('User does not have permission to edit this event');
      return;
    }
    if (formData.endTime < formData.startTime) {
      setTimeError("End time must be after start time.");
      return;
    }

    const eventToSave = { ...formData };

    // Include recurrence data if it exists
    if (returnedRecurrenceInfo) {
      eventToSave.RecurrenceData = returnedRecurrenceInfo.recurrenceData;
    } else if (recurrenceData && editRecurrenceEvent) {
      eventToSave.RecurrenceData = recurrenceData;
    }

    try {
      let Item: any = {
        Title: eventToSave.title,
        Location: eventToSave.locations,
        Description: eventToSave.description,
        Category: eventToSave.category,
        fAllDayEvent: isAllDay,
      };

      // Format dates properly based on whether it's an all-day event or not
      if (isAllDay) {
        // For SharePoint all-day events, both start and end dates should be the same day
        // Extract just the date portion without any timezone conversion
        const dateOnly = moment(eventToSave.startTime).format('YYYY-MM-DD');

        // Set both start and end time to the same date at 00:00:00 and WITHOUT 'Z' suffix
        Item.EventDate = `${dateOnly}T00:00:00`;
        Item.EndDate = `${dateOnly}T23:59:59`;  // End of the same day

        // Make sure the all-day flag is set
        Item.fAllDayEvent = true;
      } else {
        // For regular events with specific times
        Item.EventDate = moment(eventToSave.startTime, 'YYYY-MM-DD HH:mm:ss').format('YYYY-MM-DDTHH:mm:ss') + 'Z';
        Item.EndDate = moment(eventToSave.endTime, 'YYYY-MM-DD HH:mm:ss').format('YYYY-MM-DDTHH:mm:ss') + 'Z';
      }

      // Include recurrence data if it exists
      if (eventToSave.RecurrenceData) {
        Item.RecurrenceData = eventToSave.RecurrenceData;
      }

      if (props.isNew) {
        await sp.web.lists.getById(props.MarketingCalendarId).items.add(Item).then((response: any) => {
          console.log('Item Added successfully:', response);
          props.onSave(eventToSave);
        });
      } else {
        await sp.web.lists.getById(props.MarketingCalendarId).items.getById(parseInt(eventToSave.id)).update(Item).then((response: any) => {
          console.log('Item updated successfully:', response);
          props.onSave(eventToSave);
        })
      }
    } catch (error) {
      console.error('Error creating event:', error);
    }
  };

  const handleDelete = (): void => {
    if (formData.id && !props.isNew) {
      props.onDelete(formData.id);
    }
  };

  // Function to disable past dates or dates before start date
  const onRenderStartDatePickerDay = (date: Date): JSX.Element => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return <div>{date.getDate()}</div>;
  };

  const onRenderEndDatePickerDay = (date: Date): JSX.Element => {
    const startDate = new Date(formData.startTime);
    startDate.setHours(0, 0, 0, 0);

    // For end date, disable dates before the start date
    const isBeforeStartDate = date < startDate;

    return (
      <div style={isBeforeStartDate ? { color: '#d0d0d0', fontWeight: 'normal' } : {}}>
        {date.getDate()}
      </div>
    );
  };

  // Custom function to handle time change safely
  const handleTimeChange = (field: 'startTime' | 'endTime', newTime: Date | null) => {
    if (!newTime) return;

    try {
      const currentDate = formData[field];
      const updatedDate = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth(),
        currentDate.getDate(),
        newTime.getHours(),
        newTime.getMinutes(),
        0
      );

      handleInputChange(field, updatedDate);

      // If updating start time, check if we need to adjust end time
      if (field === 'startTime') {
        const currentEndTime = formData.endTime;

        // If new start time is later than end time, adjust end time
        if (updatedDate >= currentEndTime) {
          const newEndTime = new Date(updatedDate.getTime() + 60 * 60 * 1000); // 1 hour later
          handleInputChange('endTime', newEndTime);
        }
      }
    } catch (error) {
      console.error(`Error updating ${field}:`, error);
    }
  };

  const returnRecurrenceInfo = (startDate: Date, endDate: Date, recurrenceData: string) => {
    const returnedRecurrenceInfo = {
      recurrenceData: recurrenceData,
      eventDate: startDate,
      endDate: endDate,
    };
    if (props.isNew || props.event?.RecurrenceData == null) {
      setFormData((prevData) => ({
        ...prevData,
        RecurrenceData: returnedRecurrenceInfo?.recurrenceData,
      }));
      setReturnedRecurrenceInfo(returnedRecurrenceInfo);
    } else {
      setTempRecurrenceData(returnedRecurrenceInfo);
    }

    console.log(returnedRecurrenceInfo);
  };

  const handleRecurrenceCheck = (
    ev: React.FormEvent<HTMLElement | HTMLInputElement>,
    recurChecked: boolean
  ) => {
    ev.preventDefault();
    setShowRecurrenceSeriesInfo(recurChecked);
    setNewRecurrenceEvent(recurChecked);
  };

  const getRecurrenceType = (recurrenceData: any) => {
    if (!recurrenceData) return '';
    if (recurrenceData.indexOf('<daily') !== -1) return 'daily';
    if (recurrenceData.indexOf('<weekly') !== -1) return 'weekly';
    if (recurrenceData.indexOf('<monthly') !== -1) return 'monthly';
    if (recurrenceData.indexOf('<yearly') !== -1) return 'yearly';
    return '';
  };

  const saveRecurrenceChanges = () => {
    setRecurrenceData(tempRecurrenceData?.recurrenceData);
    setReturnedRecurrenceInfo(tempRecurrenceData);
    setFormData((prevData) => ({
      ...prevData,
      RecurrenceData: tempRecurrenceData?.recurrenceData,
    }));
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

  const getRecurrenceRule = (recurrenceData: any) => {
    return getRecurrenceType(recurrenceData); // Same as type for this use case
  };

  const renderDescriptionField = () => {
    if (props?.CalendarTitle === 'Company Calendar') {
      return (
        <div>
          <Label>Description</Label>
          <RichText
            value={formData.description}
            onChange={handleRichTextChange}
            isEditMode={true}
          />
          {errors.description && (
            <Text variant="small" style={{ color: '#d13438' }}>
              {errors.description}
            </Text>
          )}
        </div>
      );
    } else {
      return (
        <TextField
          label="Description"
          multiline
          rows={3}
          value={formData.description}
          onChange={(_, value) => handleInputChange('description', value)}
        />
      );
    }
  };

  const renderFooterContent = () => (
    <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
      <PrimaryButton onClick={handleSave} text="Save" />
      {!props.isNew && <DefaultButton onClick={handleDelete} text="Delete" />}
      <DefaultButton onClick={props.onCancel} text="Cancel" />
    </Stack>
  );

  return (
    <Panel
      isOpen={true}
      onDismiss={props.onCancel}
      headerText={props.isNew ? "Add Event" : "Edit Event"}
      type={PanelType.medium}
      onRenderFooterContent={renderFooterContent}
      isFooterAtBottom={true}
    >
      <Stack tokens={stackTokens} styles={stackStyles}>
        <TextField
          label="Title"
          value={formData.title}
          onChange={(_, value) => handleInputChange('title', value)}
          required
        />

        <TextField
          label="Location"
          value={formData.locations}
          onChange={(_, value) => handleInputChange('locations', value)}
        />

        {timeError && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline={false}
            dismissButtonAriaLabel="Close"
          >
            {timeError}
          </MessageBar>
        )}

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
          <Stack.Item grow={1}>
            {!isAllDay && (
              <TimePicker
                label="Start Time"
                value={formData.startTime}
                onChange={(_, time) => {
                  if (time) {
                    handleTimeChange('startTime', time);
                  }
                }}
                increments={15}
              />
            )}
          </Stack.Item>
        </Stack>

        <Stack horizontal tokens={stackTokens}>
          <Stack.Item grow={1}>
            <DatePicker
              label="End Date"
              value={formData.endTime}
              onSelectDate={(date) => {
                if (date) {
                  // Only allow dates on or after the start date
                  if (date >= new Date(formData.startTime.setHours(0, 0, 0, 0))) {
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
          <Stack.Item grow={1}>
            {!isAllDay && (
              <TimePicker
                label="End Time"
                value={formData.endTime}
                onChange={(_, time) => {
                  if (time) {
                    handleTimeChange('endTime', time);
                  }
                }}
                increments={15}
              />
            )}
          </Stack.Item>
        </Stack>

        <Toggle
          label="All Day Event"
          checked={isAllDay}
          onChange={(_, checked) => {
            setIsAllDay(!!checked);
            if (checked) {
              // Set start to 12:00 AM and end to 11:59 PM of the selected dates
              const start = new Date(formData.startTime);
              start.setHours(0, 0, 0, 0);
              const end = new Date(formData.endTime);
              end.setHours(23, 59, 0, 0);
              handleInputChange('startTime', start);
              handleInputChange('endTime', end);
            }
          }}
        />

        {renderDescriptionField()}

        <Dropdown
          label="Category"
          options={categoryOptions}
          selectedKey={formData.category}
          onChange={(_, option) => option && handleInputChange('category', option.key)}
        />

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
              <>{parseRecurrenceToString(recurrenceData)}</>
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
              siteUrl={ props?.siteUrl != undefined ? props?.siteUrl : props.Context.pageContext.web.absoluteUrl}
              recurrenceData={!isEditingRecurrence ? recurrenceData : ''}
              removeRecurrence={true}  // Enable editing of recurrence
              selectedRecurrenceRule={!isEditingRecurrence ? getRecurrenceRule(returnedRecurrenceInfo?.recurrenceData) : ''}
              selectedKey={!isEditingRecurrence ? getRecurrenceType(returnedRecurrenceInfo?.recurrenceData) : ''}
              display={true}  // Provide a valid value for 'display'
              DueDate={formData.endTime}  // Provide a valid value for 'DueDate'
            />

            {isEditingRecurrence && !props.isNew && (
              <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                <PrimaryButton text="Update Recurrence" onClick={saveRecurrenceChanges} />
                <DefaultButton text="Cancel" onClick={cancelRecurrenceChanges} />
                <a href="http://" target="_blank" rel="noopener noreferrer">OOTB</a>
              </Stack>
            )}
          </Stack>
        )}
      </Stack>
    </Panel>
  );
};

export default EventForm;