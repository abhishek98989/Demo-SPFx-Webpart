import * as React from 'react';
import { useState, useEffect } from 'react';
import { IModernCalendarProps } from './IModernCalendarProps';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Calendar, momentLocalizer } from 'react-big-calendar';
import moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import EventForm from './EventForm';
export interface ICalendarEvent {
    id: string;
    title: string;
    locations: string;
    startTime: Date;
    endTime: Date;
    description: string;
    attendees: any[];
    category: string;
    resources: string;
    freeBusy: string;
    checkDoubleBooking: boolean;
    modified: Date;
    created: Date;
    createdBy: any;
    modifiedBy: any;
}


const localizer = momentLocalizer(moment);

export default function ModernCalendar(props: any) {
    const [events, setEvents] = useState<ICalendarEvent[]>([]);
    const [showModal, setShowModal] = useState<boolean>(false);
    const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(null);
    const [isNewEvent, setIsNewEvent] = useState<boolean>(true);
    const sp = spfi().using(SPFx(props?.Context));
    useEffect(() => {
        if (props.MarketingCalendarId) {
            loadEvents();
        }
    }, [props.MarketingCalendarId]);

    const loadEvents = async (): Promise<void> => {
        if (!props.MarketingCalendarId) {
            console.error('List name not provided');
            return;
        }

        try {
            // Using PnP JS to get list items with the correct column names
            const items = await sp.web.lists.getById(props?.MarketingCalendarId)
                .items
                .select('Id,Title,Location,EventDate,EndDate,Description,ParticipantsPicker/Id,Category,FreeBusy,Overbook,Modified,Created,Author/Title,Editor/Title')
                .expand('Author,Editor,ParticipantsPicker').top(5000)()

            const calendarEvents: ICalendarEvent[] = items.map((item: any) => {
                return {
                    id: item.Id.toString(),
                    title: item.Title || '',
                    locations: item.Location || '',
                    startTime: new Date(item.EventDate),
                    endTime: new Date(item.EndDate),
                    description: item.Description || '',
                    attendees: item.ParticipantsPicker || [],
                    category: item.Category || '',
                    resources: item.Resources || '',
                    freeBusy: item.FreeBusy || '',
                    checkDoubleBooking: item.Overbook || false,
                    modified: new Date(item.Modified),
                    created: new Date(item.Created),
                    createdBy: item.Author ? item.Author.Title : '',
                    modifiedBy: item.Editor ? item.Editor.Title : ''
                };
            });

            setEvents(calendarEvents);
        } catch (error) {
            console.error('Error loading events:', error);
        }
    };

    const handleSelectEvent = (event: ICalendarEvent) => {
        setSelectedEvent(event);
        setIsNewEvent(false);
        setShowModal(true);
    };

    const handleSelectSlot = ({ start, end }: { start: Date; end: Date }) => {
        const newEvent: ICalendarEvent = {
            id: '',
            title: '',
            locations: '',
            startTime: start,
            endTime: end,
            description: '',
            attendees: [],
            category: '',
            resources: '',
            freeBusy: 'Busy',
            checkDoubleBooking: false,
            modified: new Date(),
            created: new Date(),
            createdBy: null,
            modifiedBy: null
        };
        setSelectedEvent(newEvent);
        setIsNewEvent(true);
        setShowModal(true);
    };

    const saveEvent = (event: ICalendarEvent) => {
        if (isNewEvent) {
            createEvent(event);
        } else {
            updateEvent(event);
        }
        setShowModal(false);
    };

    const createEvent = async (event: ICalendarEvent) => {
        try {
            const addItem = {
                Title: event.title,
                Location: event.locations,
                EventDate: event.startTime.toISOString(),
                EndDate: event.endTime.toISOString(),
                Description: event.description,
                Category: event.category,
                Resources: event.resources,
                FreeBusy: event.freeBusy,
                Overbook: event.checkDoubleBooking
            };

            await sp.web.lists.getById(props.MarketingCalendarId).items.add(addItem);
            loadEvents();
        } catch (error) {
            console.error('Error creating event:', error);
        }
    };

    const updateEvent = async (event: ICalendarEvent) => {
        try {
            const updateItem = {
                Title: event.title,
                Location: event.locations,
                EventDate: event.startTime.toISOString(),
                EndDate: event.endTime.toISOString(),
                Description: event.description,
                Category: event.category,
                Resources: event.resources,
                FreeBusy: event.freeBusy,
                Overbook: event.checkDoubleBooking
            };

            await sp.web.lists.getById(props.MarketingCalendarId).items.getById(parseInt(event.id)).update(updateItem);
            loadEvents();
        } catch (error) {
            console.error('Error updating event:', error);
        }
    };

    const deleteEvent = async (eventId: string) => {
        try {
            await sp.web.lists.getById(props.MarketingCalendarId).items.getById(parseInt(eventId)).recycle();
            loadEvents();
            setShowModal(false);
        } catch (error) {
            console.error('Error deleting event:', error);
        }
    };

    return (
        <div className={'modernCalendar'}>
            <div className={'container'}>
                <div className={'row'}>
                    <div className={'column'}>
                        <h2>Calendar for {'listName'}</h2>
                        <div className={'calendarContainer'}>
                            <Calendar
                                localizer={localizer}
                                events={events.map(event => ({
                                    ...event,
                                    start: event.startTime,
                                    end: event.endTime
                                }))}
                                startAccessor="startTime"
                                endAccessor="endTime"
                                titleAccessor="title"
                                style={{ height: 600 }}
                                selectable
                                onSelectEvent={handleSelectEvent}
                                onSelectSlot={handleSelectSlot}
                                views={['month', 'week', 'day', 'agenda']}
                            />
                        </div>
                    </div>
                </div>
            </div>
            {showModal && (
                <EventForm
                    event={selectedEvent}
                    isNew={isNewEvent}
                    onSave={saveEvent}
                    onDelete={deleteEvent}
                    onCancel={() => setShowModal(false)}
                />
            )}
        </div>
    );
}