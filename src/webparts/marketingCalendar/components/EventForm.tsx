import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import {
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  IStackStyles,
  Label
} from '@fluentui/react';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { TimePicker } from '@fluentui/react/lib/TimePicker';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { ICalendarEvent } from './CalendarEvent';
import { EventRecurrenceInfo } from '../../../globalCommon/EventRecurrenceControls/EventRecurrenceInfo/EventRecurrenceInfo';
import { spfi, SPFx } from '@pnp/sp';
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
  MarketingCalendarId: any;
  onSave: (event: ICalendarEvent) => void;
  onDelete: (eventId: string) => void;
  onCancel: () => void;
}

const EventForm: React.FC<IEventFormProps> = (props) => {
  const sp = spfi().using(SPFx(props?.Context));
  const [isAllDay, setIsAllDay] = useState(false);
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
  const [isEditingRecurrence, setIsEditingRecurrence] = useState<boolean>(false);
  const [tempRecurrenceData, setTempRecurrenceData] = useState<any | undefined>();
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

  const handleSave = async () => {
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
        EventDate: eventToSave.startTime.toISOString(),
        EndDate: eventToSave.endTime.toISOString(),
        Description: eventToSave.description,
        Category: eventToSave.category,
        fAllDayEvent: isAllDay,
      };
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

  const categoryOptions: IDropdownOption[] = [
    { key: 'Meeting', text: 'Meeting' },
    { key: 'Work hours', text: 'Work hours' },
    { key: 'Gifts', text: 'Gifts' },
    { key: 'Business', text: 'Business' },
    { key: 'Birthday', text: 'Birthday' },
    { key: 'Get-together', text: 'Get-together' },
    { key: 'Anniversary', text: 'Anniversary' }
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
    if (props.isNew) {
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
    setRecurrenceData(tempRecurrenceData);
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
  const parseRecurrenceToString = (recurrenceData?: string): string => {
    if (!recurrenceData) return 'No recurrence';

    // This is a placeholder - you'll need to implement actual parsing based on your recurrence format
    // Example implementation (modify according to your recurrence data format):
    if (recurrenceData.includes('DAILY')) {
      return 'Occurs every day';
    } else if (recurrenceData.includes('WEEKLY')) {
      return 'Occurs weekly';
    } else if (recurrenceData.includes('MONTHLY')) {
      return 'Occurs monthly';
    } else if (recurrenceData.includes('YEARLY')) {
      return 'Occurs yearly';
    }

    return 'Custom recurrence pattern';
  };
  return (
    <Panel
      isOpen={true}
      onDismiss={props.onCancel}
      headerText={props.isNew ? "Add Event" : "Edit Event"}
      type={PanelType.medium}
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
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            {!isAllDay && (
              <TimePicker
                label="Start Time"
                value={formData.startTime}
                onChange={(time) => {
                  if (time) {
                    handleInputChange('startTime', time);
                  }
                }}
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
                  const newDate = new Date(formData.endTime);
                  newDate.setFullYear(date.getFullYear(), date.getMonth(), date.getDate());
                  handleInputChange('endTime', newDate);
                }
              }}
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            {!isAllDay && (
              <TimePicker
                label="End Time"
                value={formData.endTime}
                onChange={(time) => {
                  if (time) {
                    handleInputChange('endTime', time);
                  }
                }}
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
              < >{parseRecurrenceToString(recurrenceData)}</>
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
              siteUrl={props.Context.pageContext.web.absoluteUrl}
              recurrenceData={isEditingRecurrence ? tempRecurrenceData : recurrenceData}
              removeRecurrence={true}  // Enable editing of recurrence
              selectedRecurrenceRule={getRecurrenceRule(isEditingRecurrence ? tempRecurrenceData : recurrenceData)}
              selectedKey={getRecurrenceType(isEditingRecurrence ? tempRecurrenceData : recurrenceData)}
              display={true}  // Provide a valid value for 'display'
              DueDate={formData.endTime}  // Provide a valid value for 'DueDate'
            />

            {isEditingRecurrence && !props.isNew && (
              <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                <PrimaryButton text="Update Recurrence" onClick={saveRecurrenceChanges} />
                <DefaultButton text="Cancel" onClick={cancelRecurrenceChanges} />
              </Stack>
            )}
          </Stack>
        )}


        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
          <PrimaryButton onClick={handleSave} text="Save" />
          {!props.isNew && (
            <DefaultButton onClick={handleDelete} text="Delete" />
          )}
          <DefaultButton onClick={props.onCancel} text="Cancel" />
        </Stack>
      </Stack>
    </Panel>
  );
};

export default EventForm;