import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { TimePicker } from '@fluentui/react/lib/TimePicker';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { ICalendarEvent } from './CalendarEvent';
import { EventRecurrenceInfo } from '../../../globalCommon/EventRecurrenceControls/EventRecurrenceInfo/EventRecurrenceInfo';

interface IEventFormProps {
  event: ICalendarEvent | null;
  isNew: boolean;
  Context:any;
  onSave: (event: ICalendarEvent) => void;
  onDelete: (eventId: string) => void;
  onCancel: () => void;
}

const EventForm: React.FC<IEventFormProps> = (props) => {
  const [formData, setFormData] = useState<ICalendarEvent>({
    id: '',
    title: '',
    locations: '',
    startTime: new Date(),
    endTime: new Date(),
    description: '',
    attendees: [],
    category: '',
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
    const [newRecurrenceEvent, setNewRecurrenceEvent] = React.useState(false);
    const [editRecurrenceEvent, setEditRecurrenceEvent] = React.useState(false);
    const [returnedRecurrenceInfo, setReturnedRecurrenceInfo] =
          React.useState<{ recurrenceData: string; eventDate: Date; endDate: Date } | null>(null);
    const [recurrenceData, setRecurrenceData] = React.useState(null);
    useEffect(() => {
      if (props.event) {
        setFormData({ ...props.event });
        // Check if this is a recurring event by looking for RecurrenceData
        if (props.event.RecurrenceData) {
          setShowRecurrenceSeriesInfo(true);
          setRecurrenceData(props.event.RecurrenceData);
          setNewRecurrenceEvent(false);  // Not a new recurrence, editing existing
          setEditRecurrenceEvent(true);  // Flag as editing recurrence
        }
      }
    }, [props.event]);

  const handleInputChange = (field: keyof ICalendarEvent, value: any): void => {
    setFormData(prevData => ({
      ...prevData,
      [field]: value
    }));
  };

  const handleSave = (): void => {
    const eventToSave = { ...formData };
    
    // Include recurrence data if it exists
    if (returnedRecurrenceInfo) {
      eventToSave.RecurrenceData = returnedRecurrenceInfo.recurrenceData;
    } else if (recurrenceData && editRecurrenceEvent) {
      eventToSave.RecurrenceData = recurrenceData;
    }
    
    props.onSave(eventToSave);
  };
  const handleDelete = (): void => {
    if (formData.id && !props.isNew) {
      props.onDelete(formData.id);
    }
  };

  const categoryOptions: IDropdownOption[] = [
    { key: 'Meeting', text: 'Meeting' },
    { key: 'Work', text: 'Work' },
    { key: 'Personal', text: 'Personal' },
    { key: 'Holiday', text: 'Holiday' },
    { key: 'Travel', text: 'Travel' }
  ];

  const freeBusyOptions: IDropdownOption[] = [
    { key: 'Free', text: 'Free' },
    { key: 'Tentative', text: 'Tentative' },
    { key: 'Busy', text: 'Busy' },
    { key: 'Out of Office', text: 'Out of Office' }
  ];

  const dialogContentProps = {
    type: DialogType.normal,
    title: props.isNew ? 'Add New Event' : 'Edit Event',
  };
  const returnRecurrenceInfo = (startDate: Date, endDate: Date, recurrenceData: string) => {
    const returnedRecurrenceInfo = {
      recurrenceData: recurrenceData,
      eventDate: startDate,
      endDate: endDate,
    };
    setReturnedRecurrenceInfo(returnedRecurrenceInfo);
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
  const getRecurrenceType = (recurrenceData:any) => {
    if (!recurrenceData) return '';
    if (recurrenceData.indexOf('<daily') !== -1) return 'daily';
    if (recurrenceData.indexOf('<weekly') !== -1) return 'weekly';
    if (recurrenceData.indexOf('<monthly') !== -1) return 'monthly';
    if (recurrenceData.indexOf('<yearly') !== -1) return 'yearly';
    return '';
  };
  
  const getRecurrenceRule = (recurrenceData:any) => {
    return getRecurrenceType(recurrenceData); // Same as type for this use case
  };

  return (
    <Dialog
      hidden={false}
      dialogContentProps={dialogContentProps}
      minWidth={500}
    >
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
      <div style={{ display: 'flex' }}>
        <div style={{ marginRight: '10px', width: '50%' }}>
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
          />
        </div>
        <div style={{ width: '50%' }}>
          <TimePicker
            label="Start Time"
            value={formData.startTime}
            onChange={(time) => {
              if (time) {
                handleInputChange('startTime', time);
              }
            }}
          />
        </div>
      </div>
      <div style={{ display: 'flex' }}>
        <div style={{ marginRight: '10px', width: '50%' }}>
          <DatePicker
            label="End Date"
            value={formData.endTime}
            onSelectDate={(date) => {
              if (date) {
                const newDate = new Date(formData.endTime);
                newDate.setFullYear(date.getFullYear(), date.getMonth(), date.getDate());
                handleInputChange('endTime', newDate);
              }
            }}
          />
        </div>
        <div style={{ width: '50%' }}>
          <TimePicker
            label="End Time"
            value={formData.endTime}
            onChange={(time) => {
              if (time) {
                handleInputChange('endTime', time);
              }
            }}
          />
        </div>
      </div>
      <TextField
        label="Description"
        multiline
        rows={3}
        value={formData.description}
        onChange={(_, value) => handleInputChange('description', value)}
      />
      <Dropdown
        label="Category"
        options={categoryOptions}
        selectedKey={formData.category}
        onChange={(_, option) => option && handleInputChange('category', option.key)}
      />
      <TextField
        label="Resources"
        value={formData.resources}
        onChange={(_, value) => handleInputChange('resources', value)}
      />
      <Dropdown
        label="Free/Busy"
        options={freeBusyOptions}
        selectedKey={formData.freeBusy}
        onChange={(_, option) => option && handleInputChange('freeBusy', option.key)}
      />
       {<div>
            {showRecurrence && (
              <div
                className="bdr-radius"
                style={{
                  display: "inline-block",
                  verticalAlign: "top",
                  width: "200px"
                }}
              >
                <Toggle
                  className="rounded-pill"
                  defaultChecked={false}
                  checked={showRecurrenceSeriesInfo}
                  inlineLabel
                  label="Recurrence"
                  onChange={handleRecurrenceCheck}
                  disabled={IsDisableField}
                />
              </div>
            )}
          {showRecurrenceSeriesInfo && (
  <EventRecurrenceInfo
    context={props.Context}
    startDate={formData.startTime}
    endDate={formData.endTime}
    returnRecurrenceInfo={returnRecurrenceInfo}
    siteUrl={props.Context.pageContext.web.absoluteUrl}
    recurrenceData={recurrenceData}
    removeRecurrence={true}  // Enable editing of recurrence
    selectedRecurrenceRule={getRecurrenceRule(recurrenceData)}
    selectedKey={getRecurrenceType(recurrenceData)}
    display={true}  // Provide a valid value for 'display'
    DueDate={formData.endTime}  // Provide a valid value for 'DueDate'
  />
)}
          </div>
          }
      <Toggle
        label="Check Double Booking"
        checked={formData.checkDoubleBooking}
        onChange={(_, checked) => handleInputChange('checkDoubleBooking', checked)}
      />
      <DialogFooter>
        <PrimaryButton onClick={handleSave} text="Save" />
        {!props.isNew && (
          <DefaultButton onClick={handleDelete} text="Delete" />
        )}
        <DefaultButton onClick={props.onCancel} text="Cancel" />
      </DialogFooter>
    </Dialog>
  );
};

export default EventForm;