import * as React from 'react';
import { useState, useEffect } from 'react';
import { IModernCalendarProps } from './IModernCalendarProps';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import { parseString } from 'xml2js';
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
  RecurrenceData: any;
  modifiedBy: any;
}

let AllGeneratedevents: any = [];
const localizer = momentLocalizer(moment);

export default function ModernCalendar(props: any) {
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [showModal, setShowModal] = useState<boolean>(false);
  const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(null);
  const [isNewEvent, setIsNewEvent] = useState<boolean>(true);
  const [currentCalendarDate, setCurrentCalendarDate] = useState<Date>(new Date());
  const [currentCalendarView, setCurrentCalendarView] = useState<String>('month');
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
        .select('Id,Title,Location,EventDate,RecurrenceData,fRecurrence,fAllDayEvent,EndDate,Description,ParticipantsPicker/Id,Category,FreeBusy,Overbook,Modified,Created,Author/Title,Editor/Title')
        .expand('Author,Editor,ParticipantsPicker').top(5000)()
      const NonRecurrenceData = items.filter((item) => item?.RecurrenceData == null);
      const Recurrencedatas = items.filter((item) => item?.RecurrenceData != null && item?.RecurrenceData != 'Every 1 day(s)');

      const calendarEvents: ICalendarEvent[] = NonRecurrenceData.map((item: any) => {
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
          modifiedBy: item.Editor ? item.Editor.Title : '',
          RecurrenceData: item.RecurrenceData || null,
        };
      });
      AllGeneratedevents = [];
      AllGeneratedevents = AllGeneratedevents.concat(calendarEvents);
      for (const event of Recurrencedatas) {
        let allDates = parseRecurrence(event)
        if (allDates.length > 0) {
          AllGeneratedevents = AllGeneratedevents.concat(allDates)
        }
      }
      handleNavigate(currentCalendarDate, currentCalendarView)
      // setEvents(AllGeneratedevents);
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
      modifiedBy: null,
      RecurrenceData: ''
    };
    setSelectedEvent(newEvent);
    setIsNewEvent(true);
    setShowModal(true);
  };
  /**
 * Parse recurrence data from SharePoint calendar events
 * @param recurrenceData Event recurrence data from SharePoint
 * @returns Array of recurring event instances
 */

  function eventDataForBinding(eventDetails: any, currentDate: Date) {
    return {
      ...eventDetails,
      id: eventDetails.Id,
      EndDate: new Date(currentDate).toISOString(),
      EventDate: new Date(currentDate).toISOString(),
      title: eventDetails.Title,
      start: new Date(currentDate),
      end: new Date(currentDate),
      startTime: new Date(currentDate),
      endTime: new Date(currentDate),
      RecurrenceData: eventDetails.RecurrenceData
    };
  }

  /**
   * Calculate the next date in a recurrence pattern
   * @param rule Recurrence rule
   * @param firstDayOfWeek First day of week setting
   * @param currentDate Current date being processed
   * @param dates Array of dates already processed
   * @param endDate End date for recurrence
   * @param allEvents Array of event objects
   * @param eventDetails Original event details
   * @param eventStartDate Original event start date to prevent events before this date
   * @returns String 'break' if processing should stop, or empty string
   */
  function parseRecurrence(recurrenceData: any) {
    const dates: Date[] = [];
    const allEvents: any[] = [];

    try {
      // Add global safety counter to prevent excessive processing
      let globalSafetyCounter = 0;
      const GLOBAL_MAX_ITERATIONS = 10000;

      parseString(recurrenceData?.RecurrenceData, (err: any, result: any) => {
        if (err || !result || !result.recurrence) {
          console.error('Error parsing XML:', err);
          return [];
        }

        const { recurrence } = result;
        const rule = recurrence?.rule?.[0];
        const firstDayOfWeek = rule?.firstDayOfWeek || 'su';
        const startDate = new Date(recurrenceData?.EventDate);
        let repeatInstance = rule?.repeatInstances ? Number(rule.repeatInstances[0]) : 0;

        // Determine end date for recurrence
        let windowEndDate: Date;

        if (rule?.repeatForever && rule?.repeatForever[0] === 'FALSE') {
          if (rule?.windowEnd === undefined) {
            // If no window end specified, create a reasonable default (1000 days)
            const createEndDate = new Date(recurrenceData?.EndDate);
            createEndDate.setHours(0, 0, 0, 0);
            createEndDate.setDate(createEndDate.getDate() + 1000);
            windowEndDate = createEndDate;
          } else {
            windowEndDate = new Date(rule.windowEnd[0]);
            windowEndDate.setHours(0, 0, 0, 0);
          }
        } else if (repeatInstance > 0) {
          // Calculate end date based on repeat instances
          const repeatInstanceEndDate = new Date(recurrenceData?.EventDate);
          repeatInstanceEndDate.setHours(0, 0, 0, 0);

          const { repeat } = rule;
          const repeatType = Object.keys(repeat[0])[0];

          switch (repeatType) {
            case 'daily':
              repeatInstanceEndDate.setDate(repeatInstanceEndDate.getDate() + repeatInstance);
              break;
            case 'weekly':
              repeatInstanceEndDate.setDate(repeatInstanceEndDate.getDate() + (7 * repeatInstance));
              break;
            case 'monthly':
              repeatInstanceEndDate.setMonth(repeatInstanceEndDate.getMonth() + repeatInstance);
              break;
            case 'yearly':
              repeatInstanceEndDate.setFullYear(repeatInstanceEndDate.getFullYear() + repeatInstance);
              break;
          }

          windowEndDate = repeatInstanceEndDate;
        } else {
          // Use the end date from recurrence data
          windowEndDate = new Date(recurrenceData?.EndDate);
          windowEndDate.setHours(0, 0, 0, 0);
        }

        // Generate recurring events
        let currentDate = new Date(startDate);

        // Important: Check if current date is before the actual event start date
        // This fixes the bug where events are created before the specified start date
        const eventStartDate = new Date(recurrenceData?.EventDate);
        eventStartDate.setHours(0, 0, 0, 0);

        // FIX: Move the counter check into the loop condition directly
        // This ensures we only continue if we haven't hit our target repeats
        // and we're still within the window end date
        while (
          (repeatInstance === 0 || allEvents.length < repeatInstance) &&
          (allEvents.length === 0 || new Date(dates[dates.length - 1]).setHours(0, 0, 0, 0) < windowEndDate.getTime()) &&
          globalSafetyCounter < GLOBAL_MAX_ITERATIONS
        ) {
          // FIX: Check the return value to properly handle breaking
          const result = calculateNextDate(
            rule,
            firstDayOfWeek,
            currentDate,
            dates,
            windowEndDate,
            allEvents,
            recurrenceData,
            eventStartDate,
            repeatInstance // Pass required repeat count
          );

          if (result === 'break' || allEvents.length >= repeatInstance && repeatInstance > 0) {
            break;
          }

          // FIX: Advance current date to prevent infinite loops
          // If no dates were added in this iteration, advance by one day to prevent infinite loop
          if (result === 'no-dates-added') {
            currentDate.setDate(currentDate.getDate() + 1);
          }

          // Increment safety counter
          globalSafetyCounter++;
        }

        // Check if we hit the global safety limit
        if (globalSafetyCounter >= GLOBAL_MAX_ITERATIONS) {
          console.warn("Global safety limit reached in parseRecurrence: possible infinite loop prevented");
        }
      });
    } catch (error) {
      console.error("Parsing error", error);
    }

    return allEvents;
  }

  /**
   * Calculate the next date in a recurrence pattern
   * @param rule Recurrence rule
   * @param firstDayOfWeek First day of week setting
   * @param currentDate Current date being processed
   * @param dates Array of dates already processed
   * @param endDate End date for recurrence
   * @param allEvents Array of event objects
   * @param eventDetails Original event details
   * @param eventStartDate Original event start date to prevent events before this date
   * @param repeatInstance Maximum number of instances to generate
   * @returns String 'break' if processing should stop, 'no-dates-added' if no dates added, or empty string
   */
  function calculateNextDate(
    rule: any,
    firstDayOfWeek: string,
    currentDate: Date,
    dates: Date[],
    endDate: Date,
    allEvents: any[],
    eventDetails: any,
    eventStartDate: Date,
    repeatInstance: number
  ): string {
    try {
      // Safety check to prevent processing dates far beyond the end date
      if (currentDate.getTime() > endDate.getTime() + (90 * 24 * 60 * 60 * 1000)) { // 90 days grace period
        console.warn("Date is far beyond end date, stopping recurrence calculation");
        return 'break';
      }

      // FIX: Check if we've reached max repeat instances
      if (repeatInstance > 0 && allEvents.length >= repeatInstance) {
        return 'break';
      }

      const { repeat } = rule;
      if (!repeat || !repeat[0]) {
        console.error("Invalid repeat rule structure");
        return 'break';
      }

      const repeatType = Object.keys(repeat[0])[0];
      const frequency = repeat[0][repeatType][0].$;

      // FIX: Store the current length to detect if dates were added
      const initialLength = allEvents.length;

      switch (repeatType) {
        case 'daily':
          handleDailyRecurrence(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            repeatInstance,
            eventStartDate
          );
          break;

        case 'weekly':
          handleWeeklyRecurrence(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            eventStartDate
          );
          break;

        case 'monthly':
          handleMonthlyRecurrence(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            eventStartDate
          );
          break;

        case 'monthlyByDay':
          handleMonthlyByDay(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            eventStartDate
          );
          break;

        case 'yearlyByDay':
          handleYearlyByDay(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            eventStartDate
          );
          break;

        case 'yearly':
          handleYearlyRecurrence(
            frequency,
            currentDate,
            dates,
            allEvents,
            eventDetails,
            endDate,
            eventStartDate
          );
          break;

        default:
          console.warn(`Unsupported recurrence type: ${repeatType}`);
          return 'break';
      }

      // FIX: Check if any dates were added during this iteration
      if (allEvents.length === initialLength) {
        return 'no-dates-added';
      }

      // Update the current date to the last date added
      if (dates.length > 0) {
        // FIX: Move current date forward past the last added date
        currentDate.setTime(dates[dates.length - 1].getTime());
        currentDate.setDate(currentDate.getDate() + 1);
      }

    } catch (error) {
      console.error("Date calculation error", error);
      return 'break'; // Return 'break' on error to prevent further processing
    }

    return '';
  }

  /**
   * Handle daily recurrence pattern
   */
  function handleDailyRecurrence(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    repeatInstance: number,
    eventStartDate: Date
  ) {
    const dayFrequency = parseInt(frequency.dayFrequency);
    let nextDate = new Date(currentDate);

    // If using weekday pattern, handle differently
    if (frequency?.weekday === 'TRUE') {
      // Create array of weekdays
      const weekdays = [];
      let tempDate = new Date(nextDate);

      // Add safety counter to prevent infinite loops
      let safetyCounter = 0;
      const MAX_ITERATIONS = 10000; // Reasonable upper limit for iterations

      // FIX: Get enough weekdays to satisfy the repeat instance or end date
      // and check against the current allEvents length
      while ((repeatInstance === 0 || allEvents.length + weekdays.length < repeatInstance) &&
        tempDate.getTime() < windowEndDate.getTime() &&
        safetyCounter < MAX_ITERATIONS) {

        tempDate.setDate(tempDate.getDate() + 1);
        const dayOfWeek = tempDate.getDay();

        // Include only weekdays (1-5, Monday through Friday)
        if (dayOfWeek >= 1 && dayOfWeek <= 5) {
          // Only add dates on or after the start date
          if (tempDate.getTime() >= eventStartDate.getTime()) {
            weekdays.push(new Date(tempDate));

            // FIX: Break early if we've reached the target repeat count
            if (repeatInstance > 0 && allEvents.length + weekdays.length >= repeatInstance) {
              break;
            }
          }
        }

        // Increment safety counter
        safetyCounter++;
      }

      // Check if we hit the safety limit
      if (safetyCounter >= MAX_ITERATIONS) {
        console.warn("Safety limit reached in handleDailyRecurrence: possible infinite loop prevented");
      }

      // Add events for each weekday
      for (const weekday of weekdays) {
        if (weekday.getTime() <= windowEndDate.getTime()) {
          const event = eventDataForBinding(eventDetails, weekday);
          allEvents.push(event);
          dates.push(new Date(weekday));

          // FIX: Break if we've reached the repeat count
          if (repeatInstance > 0 && allEvents.length >= repeatInstance) {
            break;
          }
        }
      }
    } else {
      // Regular daily pattern
      let safetyCounter = 0;
      const MAX_ITERATIONS = 10000;

      // FIX: Check against allEvents.length directly
      while ((repeatInstance === 0 || allEvents.length < repeatInstance) &&
        safetyCounter < MAX_ITERATIONS) {
        nextDate.setDate(nextDate.getDate() + dayFrequency);

        // Only add dates on or after the start date
        if (nextDate.getTime() >= eventStartDate.getTime() && nextDate.getTime() <= windowEndDate.getTime()) {
          const event = eventDataForBinding(eventDetails, nextDate);
          allEvents.push(event);
          dates.push(new Date(nextDate));

          // FIX: No need for separate count variable
        }

        if (nextDate.getTime() > windowEndDate.getTime()) {
          break;
        }

        // Increment safety counter
        safetyCounter++;
      }

      // Check if we hit the safety limit
      if (safetyCounter >= MAX_ITERATIONS) {
        console.warn("Safety limit reached in handleDailyRecurrence: possible infinite loop prevented");
      }
    }
  }

  /**
   * Handle weekly recurrence pattern
   */
  function handleWeeklyRecurrence(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    eventStartDate: Date
  ) {
    const { weekFrequency } = frequency;
    const weekFreq = parseInt(weekFrequency);
    const daysOfWeekIndex = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];

    // Get the days of the week that are marked as TRUE
    const daysOfWeek = daysOfWeekIndex.filter(day => frequency[day] === "TRUE");

    // If no days specified, return
    if (daysOfWeek.length === 0) return;

    // Create a copy of current date to work with
    let baseDate = new Date(currentDate);
    let keepProcessing = true;

    // Add safety counter to prevent infinite loops
    let safetyCounter = 0;
    const MAX_ITERATIONS = 1000; // Reasonable upper limit for iterations

    while (keepProcessing && safetyCounter < MAX_ITERATIONS) {
      let addedThisWeek = false;

      for (const day of daysOfWeek) {
        const targetDayIndex = daysOfWeekIndex.indexOf(day);
        let targetDate = new Date(baseDate);

        // Calculate days to add to reach target day in current week
        let daysToAdd = (targetDayIndex - baseDate.getDay() + 7) % 7;
        if (daysToAdd === 0 && targetDayIndex !== baseDate.getDay()) {
          daysToAdd = 7;
        }

        targetDate.setDate(baseDate.getDate() + daysToAdd);

        // Check if this date is valid for our recurrence
        if (targetDate.getTime() >= eventStartDate.getTime() &&
          targetDate.getTime() <= windowEndDate.getTime()) {

          const event = eventDataForBinding(eventDetails, targetDate);
          allEvents.push(event);
          dates.push(new Date(targetDate));
          addedThisWeek = true;
        }

        // Stop if we've gone past the end date
        if (targetDate.getTime() > windowEndDate.getTime()) {
          keepProcessing = false;
        }
      }

      // If we've added at least one date this week, process the next week
      if (!addedThisWeek && dates.length > 0) {
        keepProcessing = false;
      }

      // Advance to next week based on week frequency
      baseDate.setDate(baseDate.getDate() + (7 * weekFreq));

      // Safety check to prevent infinite loops
      if (baseDate.getTime() > windowEndDate.getTime()) {
        keepProcessing = false;
      }

      // Increment safety counter
      safetyCounter++;
    }

    // Check if we hit the safety limit
    if (safetyCounter >= MAX_ITERATIONS) {
      console.warn("Safety limit reached in handleWeeklyRecurrence: possible infinite loop prevented");
    }
  }

  /**
   * Handle monthly recurrence pattern
   */
  function handleMonthlyRecurrence(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    eventStartDate: Date
  ) {
    const { monthFrequency, dayOfMonth } = frequency;
    const monthFreq = parseInt(monthFrequency);

    // Create a copy of the current date
    let targetDate = new Date(currentDate);
    let day = dayOfMonth || frequency?.day;

    if (!day) return;

    // Set day of month
    targetDate.setDate(parseInt(day));

    // Add safety counter
    let safetyCounter = 0;
    const MAX_ITERATIONS = 500;

    // Process monthly recurrences
    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      // Add event only if it's on or after the start date
      if (targetDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, targetDate);
        allEvents.push(event);
        dates.push(new Date(targetDate));
      }

      // Move to next month based on frequency
      targetDate.setMonth(targetDate.getMonth() + monthFreq);

      // Ensure we maintain the same day of month
      // Handle special cases like 31st of month in shorter months
      const expectedDay = parseInt(day);
      const actualMonth = targetDate.getMonth();
      targetDate.setDate(1); // Reset to 1st of month
      targetDate.setMonth(actualMonth); // Set the correct month

      // Now set the day, clamping to last day of month if needed
      const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
      targetDate.setDate(Math.min(expectedDay, lastDayOfMonth));

      // Increment safety counter
      safetyCounter++;
    }

    if (safetyCounter >= MAX_ITERATIONS) {
      console.warn("Safety limit reached in handleMonthlyRecurrence: possible infinite loop prevented");
    }
  }

  /**
   * Handle monthly by day recurrence pattern
   */
  function handleMonthlyByDay(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    eventStartDate: Date
  ) {
    const { monthFrequency, weekdayOfMonth } = frequency;
    const daysOfWeekIndex = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
    const weekMap: any = { first: 1, second: 2, third: 3, fourth: 4, last: 5 };
    const monthFreq = parseInt(monthFrequency);

    let targetDate = new Date(currentDate);

    // Add safety counter to prevent infinite loops
    let safetyCounter = 0;
    const MAX_ITERATIONS = 500; // Reasonable upper limit for iterations

    // Process recurrence until we reach the end date
    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      let specificDate: Date | null = null;

      if (frequency?.day === "TRUE") {
        // Handle specific day of the month
        specificDate = new Date(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekday === "TRUE") {
        // Handle weekdays
        // Find the nth weekday of the month
        specificDate = getNthWeekdayOfMonth(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          [1, 2, 3, 4, 5], // Monday to Friday
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekend_day === "TRUE") {
        // Handle weekend days
        specificDate = getNthWeekdayOfMonth(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          [0, 6], // Sunday and Saturday
          weekMap[weekdayOfMonth]
        );
      } else {
        // Handle specific days of the week
        const keys = Object.keys(frequency);
        const dayIndices = [];

        for (let i = 0; i < daysOfWeekIndex.length; i++) {
          const key = daysOfWeekIndex[i];
          if (keys.includes(key) && frequency[key] === "TRUE") {
            dayIndices.push(i);
          }
        }

        specificDate = getNthWeekdayOfMonth(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          dayIndices,
          weekMap[weekdayOfMonth]
        );
      }

      // Add event if date is valid
      if (specificDate &&
        specificDate.getTime() >= eventStartDate.getTime() &&
        specificDate.getTime() <= windowEndDate.getTime()) {
        const event = eventDataForBinding(eventDetails, specificDate);
        allEvents.push(event);
        dates.push(new Date(specificDate));
      }

      // Advance to next month
      targetDate.setMonth(targetDate.getMonth() + monthFreq);

      // Increment safety counter
      safetyCounter++;
    }

    // Check if we hit the safety limit
    if (safetyCounter >= MAX_ITERATIONS) {
      console.warn("Safety limit reached in handleMonthlyByDay: possible infinite loop prevented");
    }
  }

  /**
   * Get the nth occurrence of a weekday in a month
   */
  function getNthWeekdayOfMonth(
    year: number,
    month: number,
    dayIndices: number[],
    nth: number
  ): Date | null {
    if (nth === 5) { // "last" occurrence
      return getLastWeekdayOfMonth(year, month, dayIndices);
    }

    // Get first day of month
    const firstDayOfMonth = new Date(year, month, 1);

    // Find occurrences of the specified weekdays
    const occurrences = [];
    let currentDate = new Date(firstDayOfMonth);

    // Find first occurrence of each day
    for (const dayIndex of dayIndices) {
      let daysToAdd = (dayIndex - firstDayOfMonth.getDay() + 7) % 7;
      if (daysToAdd === 0 && dayIndex !== firstDayOfMonth.getDay()) {
        daysToAdd = 7;
      }

      const firstOccurrence = new Date(year, month, 1 + daysToAdd);
      occurrences.push(firstOccurrence);
    }

    // Sort by date
    occurrences.sort((a, b) => a.getTime() - b.getTime());

    // Safety check: if no occurrences found, return null
    if (occurrences.length === 0) {
      return null;
    }

    // Find the nth occurrence of any of the specified weekdays
    let count = 1;
    let lastDate = occurrences[0];

    // Add safety counter to prevent infinite loops
    let safetyCounter = 0;
    const MAX_ITERATIONS = 100; // Reasonable upper limit for iterations

    while (count < nth && safetyCounter < MAX_ITERATIONS) {
      // Find next occurrences
      const nextOccurrences = [];

      for (const date of occurrences) {
        const nextDate: any = new Date(date);
        nextDate.setDate(nextDate.getDate() + 7);

        // Only include if still in the same month
        if (nextDate.getMonth() === month) {
          nextOccurrences.push(nextDate);
        }
      }

      if (nextOccurrences.length === 0) {
        // No more occurrences this month
        break;
      }

      occurrences.length = 0;
      occurrences.push(...nextOccurrences);
      occurrences.sort((a, b) => a.getTime() - b.getTime());

      lastDate = occurrences[0];
      count++;

      // Increment safety counter
      safetyCounter++;
    }

    // Check if we hit the safety limit
    if (safetyCounter >= MAX_ITERATIONS) {
      console.warn("Safety limit reached in getNthWeekdayOfMonth: possible infinite loop prevented");
      return null;
    }

    return lastDate;
  }

  /**
   * Get the last occurrence of specified weekdays in a month
   */
  function getLastWeekdayOfMonth(
    year: number,
    month: number,
    dayIndices: number[]
  ): Date {
    // Get last day of month
    const lastDayOfMonth = new Date(year, month + 1, 0);
    const daysInMonth = lastDayOfMonth.getDate();

    // Find the last occurrence of the specified weekdays
    const lastOccurrences = [];

    for (const dayIndex of dayIndices) {
      let day = daysInMonth;

      // Start from the end of the month and work backwards
      while (day > 0) {
        const date = new Date(year, month, day);
        if (date.getDay() === dayIndex) {
          lastOccurrences.push(date);
          break;
        }
        day--;
      }
    }

    // Sort by date (descending to get the latest)
    lastOccurrences.sort((a, b) => b.getTime() - a.getTime());

    return lastOccurrences[0];
  }

  /**
   * Handle yearly recurrence pattern
   */
  function handleYearlyRecurrence(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    eventStartDate: Date
  ) {
    const { yearFrequency, month, day } = frequency;
    const yearFreq = parseInt(yearFrequency);

    // Create date for the recurring event
    let targetDate = new Date(currentDate);
    targetDate.setMonth(parseInt(month) - 1); // Month is 1-based in SharePoint
    targetDate.setDate(parseInt(day));

    // Add safety counter
    let safetyCounter = 0;
    const MAX_ITERATIONS = 200;

    // Process yearly recurrences
    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      // Add event if it's on or after the start date
      if (targetDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, targetDate);
        allEvents.push(event);
        dates.push(new Date(targetDate));
      }

      // Advance to next year
      targetDate.setFullYear(targetDate.getFullYear() + yearFreq);

      // Increment safety counter
      safetyCounter++;
    }

    if (safetyCounter >= MAX_ITERATIONS) {
      console.warn("Safety limit reached in handleYearlyRecurrence: possible infinite loop prevented");
    }
  }

  /**
   * Handle yearly by day recurrence pattern
   */
  function handleYearlyByDay(
    frequency: any,
    currentDate: Date,
    dates: Date[],
    allEvents: any[],
    eventDetails: any,
    windowEndDate: Date,
    eventStartDate: Date
  ) {
    const { yearFrequency, weekdayOfMonth, month } = frequency;
    const daysOfWeekIndex = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
    const weekMap: any = { first: 1, second: 2, third: 3, fourth: 4, last: 5 };
    const yearFreq = parseInt(yearFrequency);

    let targetYear = currentDate.getFullYear();
    const targetMonth = parseInt(month) - 1; // Month is 1-based in SharePoint

    // Add safety counter to prevent infinite loops
    let safetyCounter = 0;
    const MAX_ITERATIONS = 200; // Reasonable upper limit for iterations

    // Process yearly recurrences
    while (safetyCounter < MAX_ITERATIONS) {
      let specificDate: Date | null = null;

      if (frequency?.day === "TRUE") {
        // Handle specific day of the month
        specificDate = new Date(
          targetYear,
          targetMonth,
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekday === "TRUE") {
        // Handle weekdays
        specificDate = getNthWeekdayOfMonth(
          targetYear,
          targetMonth,
          [1, 2, 3, 4, 5], // Monday to Friday
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekend_day === "TRUE") {
        // Handle weekend days
        specificDate = getNthWeekdayOfMonth(
          targetYear,
          targetMonth,
          [0, 6], // Sunday and Saturday
          weekMap[weekdayOfMonth]
        );
      } else {
        // Handle specific days of the week
        const keys = Object.keys(frequency);
        const dayIndices = [];

        for (let i = 0; i < daysOfWeekIndex.length; i++) {
          const key = daysOfWeekIndex[i];
          if (keys.includes(key) && frequency[key] === "TRUE") {
            dayIndices.push(i);
          }
        }

        specificDate = getNthWeekdayOfMonth(
          targetYear,
          targetMonth,
          dayIndices,
          weekMap[weekdayOfMonth]
        );
      }

      // Check if we've gone past the end date or if specificDate is null
      if (!specificDate || specificDate.getTime() > windowEndDate.getTime()) {
        break;
      }

      // Add event if it's on or after the start date
      if (specificDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, specificDate);
        allEvents.push(event);
        dates.push(new Date(specificDate));
      }

      // Advance to next year
      targetYear += yearFreq;

      // Increment safety counter
      safetyCounter++;
    }

    /**
     * Handle yearly by day recurrence pattern
     */

  }

  const saveEvent = (event: ICalendarEvent) => {
    if (isNewEvent) {
      createEvent(event);
    } else {
      updateEvent(event);
    }
    setShowModal(false);
  };

  const createEvent = async (event: ICalendarEvent) => {
    loadEvents();
  };

  const updateEvent = async (event: ICalendarEvent) => {
    loadEvents();
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
  function getYearMonthFromDate(date: any) {
    const eventDate = new Date(date);
    const eventMonth = eventDate.getMonth() + 1; // Get month of the event
    const eventYear = eventDate.getFullYear(); // Get year of the event
    return { year: eventYear, month: eventMonth };
  }
  const handleNavigate = (newDate: any, newiew: any) => {
    setCurrentCalendarDate(newDate);
    setCurrentCalendarView(newiew);
    const { year: currentYear, month: currentMonth } = getYearMonthFromDate(newDate);
    const filteredData = AllGeneratedevents.filter((event: any) => {
      const startDate = getYearMonthFromDate(event.startTime);
      const endDate = getYearMonthFromDate(event.endTime);
      return (
        (startDate.year < currentYear || (startDate.year === currentYear && startDate.month <=  (currentMonth == 1 ? 12 : currentMonth - 1))) &&
        (endDate.year > currentYear || (endDate.year === currentYear && endDate.month >= (currentMonth == 12 ? 1 : currentMonth + 1)))
      );
    });
    setEvents(filteredData);
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
                onNavigate={handleNavigate}
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
          Context={props?.Context}
          MarketingCalendarId={props?.MarketingCalendarId}
          onCancel={() => setShowModal(false)}
        />
      )}
    </div>
  );
}