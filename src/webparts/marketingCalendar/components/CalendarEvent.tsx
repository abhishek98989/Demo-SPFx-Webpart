import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { IModernCalendarProps } from './IModernCalendarProps';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import { parseString } from 'xml2js';
import "@pnp/sp/items";
import { Calendar, momentLocalizer } from 'react-big-calendar';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PermissionKind } from "@pnp/sp/security";
import { format } from 'date-fns';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import EventForm from './EventForm';
import moment from 'moment';
import { useGlobalLoaderContext } from '../../../globalCommon/customLoader';
import "@pnp/sp/security";
import { Checkbox } from '@fluentui/react/lib/Checkbox';

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
  fAllDayEvent?: boolean;
  modifiedBy: any;
  Color?: string;
  FontColor?: string;
}

interface IUserPermissions {
  canAdd: boolean;
  canEdit: boolean;
  canDelete: boolean;
  canView: boolean;
}

let AllGeneratedevents: any = [];
const localizer = momentLocalizer(moment);
const NA_CATEGORY = 'N/A';

export default function ModernCalendar(props: any) {
  const { showLoader, hideLoader } = useGlobalLoaderContext();

  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [showModal, setShowModal] = useState<boolean>(false);
  const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(null);
  const [isNewEvent, setIsNewEvent] = useState<boolean>(true);
  const [currentCalendarDate, setCurrentCalendarDate] = useState<Date>(new Date());
  const [currentCalendarView, setCurrentCalendarView] = useState<String>('month');

  const sp = props?.siteUrl != undefined ? spfi(props?.siteUrl).using(SPFx(props?.Context)) : spfi().using(SPFx(props?.Context));
  const [showDatePanel, setShowDatePanel] = useState(false);
  const [selectedDateEvents, setSelectedDateEvents] = useState<ICalendarEvent[]>([]);
  const [selectedDate, setSelectedDate] = useState<Date | null>(null);

  const [userPermissions, setUserPermissions] = useState<IUserPermissions>({
    canAdd: false,
    canEdit: false,
    canDelete: false,
    canView: true
  });
  const [currentUserId, setCurrentUserId] = useState<number | null>(null);

  // ===== CATEGORIES: options-by-calendar with N/A appended =====
  type Option = { key: string; text: string };
  const [categoryOptions, setCategoryOptions] = React.useState<Option[]>([]);

  const calendarTitle = props?.CalendarTitle as string | undefined;

  useEffect(() => {
    const base: Option[] =
      calendarTitle === 'Marketing Calendar'
        ? [
            { key: 'Meeting', text: 'Meeting' },
            { key: 'RFQ', text: 'RFQ' },
            { key: 'RFP', text: 'RFP' },
            { key: 'CSP/Traditional', text: 'CSP/Traditional' },
            { key: 'DB', text: 'DB' },
            { key: 'Interview', text: 'Interview' }
          ]
        : calendarTitle === 'Marketing Calendar-Internal'
        ? [
            { key: 'PTO', text: 'PTO' },
            { key: 'Remote', text: 'Remote' },
            { key: 'Appointment', text: 'Appointment' },
            { key: 'Site Visit', text: 'Site Visit' },
            { key: 'Interview Prep', text: 'Interview Prep' },
            { key: 'Other', text: 'Other' }
          ]
        : calendarTitle === 'Company Calendar'
        ? [
            { key: 'Safety', text: 'Safety' },
            { key: 'Ops', text: 'Ops' },
            { key: 'HR', text: 'HR' },
            { key: 'Other', text: 'Other' }
          ]
        : [];

    // Always append N/A at the end (avoid dupes)
    const withNA = [...base.filter(o => o.key !== NA_CATEGORY), { key: NA_CATEGORY, text: NA_CATEGORY }];
    setCategoryOptions(withNA);
  }, [calendarTitle]);

  // ===== COLOR MAPS =====
  const isTraining = calendarTitle === 'Company Calendar';

  const categoryOptionsColor: Record<string, string> = {
    // Marketing + Internal
    'Meeting': '#3174ad',
    'RFQ': '#ffff00',
    'RFP': '#107c10',
    'CSP/Traditional': '#da3b01',
    'DB': '#c239b3',
    'Interview': '#adadad',
    'Appointment': '#0099bc',
    'Site Visit': '#00b294',
    'Remote': '#004e8c',
    'Interview Prep': '#ffaa44',
    'Other': '#605e5c',
    'PTO': '#e3008c',
    // N/A fallback (grey)
    [NA_CATEGORY]: '#605e5c',
  };
  const categoryOptionsFontColor: Record<string, string> = {
    'Meeting': '#ffffff',
    'RFQ': '#000000',
    'RFP': '#ffffff',
    'CSP/Traditional': '#ffffff',
    'DB': '#ffffff',
    'Interview': '#000000',
    'Appointment': '#ffffff',
    'Site Visit': '#000000',
    'Remote': '#ffffff',
    'Interview Prep': '#000000',
    'Other': '#ffffff',
    'PTO': '#ffffff',
    [NA_CATEGORY]: '#ffffff',
  };

  const categoryOptionsColorTraining: Record<string, string> = {
    'Safety': '#ff0000',
    'Ops': '#3174ad',
    'HR': '#107c10',
    'Other': '#ffff00',
    [NA_CATEGORY]: '#605e5c',
  };
  const categoryOptionsFontColorTraining: Record<string, string> = {
    'Safety': '#ffffff',
    'Ops': '#ffffff',
    'HR': '#ffffff',
    'Other': '#000000',
    [NA_CATEGORY]: '#ffffff',
  };

  // Build palette **only** for categories in categoryOptions (plus N/A already present)
  const categoryPalette: Record<string, { bg: string; fg: string }> = React.useMemo(() => {
    const base = isTraining ? categoryOptionsColorTraining : categoryOptionsColor;
    const baseFont = isTraining ? categoryOptionsFontColorTraining : categoryOptionsFontColor;

    const map: Record<string, { bg: string; fg: string }> = {};
    categoryOptions.forEach(opt => {
      const k = opt.key;
      map[k] = { bg: base[k] ?? '#605e5c', fg: baseFont[k] ?? '#ffffff' };
    });
    return map;
  }, [isTraining, categoryOptions]);

  // Derived, ordered list of category keys for the filter bar
  const categoryList = React.useMemo(() => categoryOptions.map(o => o.key), [categoryOptions]);

  // Filter selection
  const [selectedCategories, setSelectedCategories] = useState<Set<string>>(new Set());
  useEffect(() => {
    // Select all whenever categoryOptions change (e.g., calendar switch)
    setSelectedCategories(new Set(categoryList));
  }, [categoryList]);

  const filteredEvents = React.useMemo(() => {
    if (!selectedCategories || selectedCategories.size === 0) return [];
    return events.filter(e => selectedCategories.has(e.category ?? NA_CATEGORY));
  }, [events, selectedCategories]);

  // ===== PERMISSIONS =====
  useEffect(() => {
    if (props.MarketingCalendarId) {
      showLoader('Loading...');
      checkUserPermissions();
    }
  }, [props.MarketingCalendarId]);

  const checkUserPermissions = async (): Promise<void> => {
    if (!props?.MarketingCalendarId) {
      console.warn('No MarketingCalendarId provided for permission check');
      hideLoader();
      return;
    }

    try {
      const currentUser = await sp.web.currentUser();
      setCurrentUserId(currentUser.Id);

      const list = sp.web.lists.getById(props.MarketingCalendarId);
      let userPerms: any = null;
      const maxAttempts = 3;

      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
          userPerms = await list.getCurrentUserEffectivePermissions();
          if (userPerms) break;
        } catch {
          try {
            if (currentUser && currentUser.LoginName) {
              userPerms = await list.getUserEffectivePermissions(currentUser.LoginName);
              if (userPerms) break;
            }
          } catch {
            try {
              if (currentUser && currentUser.Id) {
                userPerms = await list.getUserEffectivePermissions(String(currentUser.Id));
                if (userPerms) break;
              }
            } catch { /* ignore */ }
          }
        }
        await new Promise(res => setTimeout(res, 200 * attempt));
      }

      if (!userPerms) {
        console.warn('Unable to determine user permissions for list after retries');
        setUserPermissions({ canAdd: false, canEdit: false, canDelete: false, canView: true });
        hideLoader();
        return;
      }

      const permissions = {
        canView: sp.web.hasPermissions(userPerms, PermissionKind.ViewListItems),
        canAdd: sp.web.hasPermissions(userPerms, PermissionKind.AddListItems),
        canEdit: sp.web.hasPermissions(userPerms, PermissionKind.EditListItems),
        canDelete: sp.web.hasPermissions(userPerms, PermissionKind.DeleteListItems)
      };

      setUserPermissions(permissions);
      if (permissions.canView) {
        await loadEvents();
      } else {
        hideLoader();
      }
    } catch (error) {
      hideLoader();
      console.error('Error checking user permissions:', error);
      setUserPermissions({ canAdd: false, canEdit: false, canDelete: false, canView: true });
    }
  };

  // ===== LOAD EVENTS =====
  const loadEvents = async (): Promise<void> => {
    if (!props.MarketingCalendarId) {
      console.error('List name not provided');
      hideLoader();
      return;
    }

    try {
      const items = await sp.web.lists.getById(props?.MarketingCalendarId)
        .items
        .select('Id,Title,Location,EventDate,RecurrenceData,fRecurrence,fAllDayEvent,EndDate,Description,ParticipantsPicker/Id,Category,FreeBusy,Overbook,Modified,Created,Author/Title,Author/Id,Editor/Title')
        .expand('Author,Editor,ParticipantsPicker').top(5000)();

      const NonRecurrenceData = items.filter((item) => item?.RecurrenceData == null);
      const Recurrencedatas = items.filter((item) => item?.RecurrenceData != null && item?.RecurrenceData != 'Every 1 day(s)');

      const isTrainingLocal = calendarTitle === 'Company Calendar';

      // Normalize non-recurring items (category -> N/A if missing)
      const calendarEvents: ICalendarEvent[] = NonRecurrenceData.map((item: any) => {
        const normCat = (item?.Category && String(item.Category).trim()) ? item.Category : NA_CATEGORY;
        const { Color, FontColor } = getCategoryColors(normCat, isTrainingLocal);
        return {
          id: item.Id.toString(),
          title: item.Title || '',
          locations: item.Location || '',
          startTime: new Date(item.EventDate?.replace('Z', '')),
          endTime: new Date(item.EndDate?.replace('Z', '')),
          description: item.Description || '',
          attendees: item.ParticipantsPicker || [],
          category: normCat,
          resources: item.Resources || '',
          freeBusy: item.FreeBusy || '',
          checkDoubleBooking: item.Overbook || false,
          modified: new Date(item.Modified),
          created: new Date(item.Created),
          createdBy: item.Author ? { Title: item.Author.Title, Id: item.Author.Id } : null,
          modifiedBy: item.Editor ? item.Editor.Title : '',
          RecurrenceData: item.RecurrenceData || null,
          fAllDayEvent: item.fAllDayEvent || false,
          Color,
          FontColor
        };
      });

      AllGeneratedevents = [];
      AllGeneratedevents = AllGeneratedevents.concat(calendarEvents);

      for (const event of Recurrencedatas) {
        const instances = parseRecurrence(event);
        if (instances.length > 0) {
          AllGeneratedevents = AllGeneratedevents.concat(instances);
        }
      }

      handleNavigate(currentCalendarDate, currentCalendarView);
    } catch (error) {
      hideLoader();
      console.error('Error loading events:', error);
    }
  };

  // ===== RECURRENCE & HELPERS =====
  function parseRecurrence(recurrenceData: any) {
    const dates: Date[] = [];
    const allEvents: any[] = [];

    try {
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

        let windowEndDate: Date;
        if (rule?.repeatForever && rule?.repeatForever[0] === 'FALSE') {
          if (rule?.windowEnd === undefined) {
            const createEndDate = new Date(recurrenceData?.EndDate);
            createEndDate.setHours(0, 0, 0, 0);
            createEndDate.setDate(createEndDate.getDate() + 1000);
            windowEndDate = createEndDate;
          } else {
            windowEndDate = new Date(rule.windowEnd[0]);
            windowEndDate.setHours(0, 0, 0, 0);
          }
        } else if (repeatInstance > 0) {
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
          windowEndDate = new Date(recurrenceData?.EndDate);
          windowEndDate.setHours(0, 0, 0, 0);
        }

        let currentDate = new Date(startDate);
        const eventStartDate = new Date(recurrenceData?.EventDate);
        eventStartDate.setHours(0, 0, 0, 0);

        while (
          (repeatInstance === 0 || allEvents.length < repeatInstance) &&
          (allEvents.length === 0 || new Date(dates[dates.length - 1]).setHours(0, 0, 0, 0) < windowEndDate.getTime()) &&
          globalSafetyCounter < GLOBAL_MAX_ITERATIONS
        ) {
          const result = calculateNextDate(
            rule,
            firstDayOfWeek,
            currentDate,
            dates,
            windowEndDate,
            allEvents,
            recurrenceData,
            eventStartDate,
            repeatInstance
          );

          if (result === 'break' || (repeatInstance > 0 && allEvents.length >= repeatInstance)) {
            break;
          }

          if (result === 'no-dates-added') {
            currentDate.setDate(currentDate.getDate() + 1);
          }

          globalSafetyCounter++;
        }

        if (globalSafetyCounter >= GLOBAL_MAX_ITERATIONS) {
          console.warn("Global safety limit reached in parseRecurrence");
        }
      });
    } catch (error) {
      console.error("Parsing error", error);
    }

    return allEvents;
  }

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
      if (currentDate.getTime() > endDate.getTime() + (90 * 24 * 60 * 60 * 1000)) {
        return 'break';
      }
      if (repeatInstance > 0 && allEvents.length >= repeatInstance) {
        return 'break';
      }

      const { repeat } = rule;
      if (!repeat || !repeat[0]) return 'break';
      const repeatType = Object.keys(repeat[0])[0];
      const frequency = repeat[0][repeatType][0].$;

      const before = allEvents.length;

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
          return 'break';
      }

      if (allEvents.length === before) {
        return 'no-dates-added';
      }

      if (dates.length > 0) {
        currentDate.setTime(dates[dates.length - 1].getTime());
        currentDate.setDate(currentDate.getDate() + 1);
      }
    } catch (error) {
      console.error("Date calculation error", error);
      return 'break';
    }
    return '';
  }

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

    if (frequency?.weekday === 'TRUE') {
      const weekdays: Date[] = [];
      let tempDate = new Date(nextDate);
      let safetyCounter = 0;
      const MAX_ITERATIONS = 10000;

      while ((repeatInstance === 0 || allEvents.length + weekdays.length < repeatInstance) &&
        tempDate.getTime() < windowEndDate.getTime() &&
        safetyCounter < MAX_ITERATIONS) {

        tempDate.setDate(tempDate.getDate() + 1);
        const dayOfWeek = tempDate.getDay();

        if (dayOfWeek >= 1 && dayOfWeek <= 5) {
          if (tempDate.getTime() >= eventStartDate.getTime()) {
            weekdays.push(new Date(tempDate));
            if (repeatInstance > 0 && allEvents.length + weekdays.length >= repeatInstance) {
              break;
            }
          }
        }
        safetyCounter++;
      }

      for (const weekday of weekdays) {
        if (weekday.getTime() <= windowEndDate.getTime()) {
          const event = eventDataForBinding(eventDetails, weekday);
          allEvents.push(event);
          dates.push(new Date(weekday));
          if (repeatInstance > 0 && allEvents.length >= repeatInstance) break;
        }
      }
    } else {
      let safetyCounter = 0;
      const MAX_ITERATIONS = 10000;

      while ((repeatInstance === 0 || allEvents.length < repeatInstance) &&
        safetyCounter < MAX_ITERATIONS) {
        nextDate.setDate(nextDate.getDate() + dayFrequency);

        if (nextDate.getTime() >= eventStartDate.getTime() && nextDate.getTime() <= windowEndDate.getTime()) {
          const event = eventDataForBinding(eventDetails, nextDate);
          allEvents.push(event);
          dates.push(new Date(nextDate));
        }

        if (nextDate.getTime() > windowEndDate.getTime()) break;
        safetyCounter++;
      }
    }
  }

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

    const daysOfWeek = daysOfWeekIndex.filter(day => frequency[day] === "TRUE");
    if (daysOfWeek.length === 0) return;

    let baseDate = new Date(currentDate);
    let keepProcessing = true;

    let safetyCounter = 0;
    const MAX_ITERATIONS = 1000;

    while (keepProcessing && safetyCounter < MAX_ITERATIONS) {
      let addedThisWeek = false;

      for (const day of daysOfWeek) {
        const targetDayIndex = daysOfWeekIndex.indexOf(day);
        let targetDate = new Date(baseDate);
        let daysToAdd = (targetDayIndex - baseDate.getDay() + 7) % 7;
        if (daysToAdd === 0 && targetDayIndex !== baseDate.getDay()) {
          daysToAdd = 7;
        }
        targetDate.setDate(baseDate.getDate() + daysToAdd);

        if (targetDate.getTime() >= eventStartDate.getTime() &&
          targetDate.getTime() <= windowEndDate.getTime()) {

          const event = eventDataForBinding(eventDetails, targetDate);
          allEvents.push(event);
          dates.push(new Date(targetDate));
          addedThisWeek = true;
        }
        if (targetDate.getTime() > windowEndDate.getTime()) {
          keepProcessing = false;
        }
      }

      if (!addedThisWeek && dates.length > 0) {
        keepProcessing = false;
      }

      baseDate.setDate(baseDate.getDate() + (7 * weekFreq));
      if (baseDate.getTime() > windowEndDate.getTime()) keepProcessing = false;
      safetyCounter++;
    }
  }

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
    let targetDate = new Date(currentDate);
    let day = dayOfMonth || frequency?.day;
    if (!day) return;

    targetDate.setDate(parseInt(day));

    let safetyCounter = 0;
    const MAX_ITERATIONS = 500;

    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      if (targetDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, targetDate);
        allEvents.push(event);
        dates.push(new Date(targetDate));
      }

      targetDate.setMonth(targetDate.getMonth() + monthFreq);
      const expectedDay = parseInt(day);
      const actualMonth = targetDate.getMonth();
      targetDate.setDate(1);
      targetDate.setMonth(actualMonth);
      const lastDayOfMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();
      targetDate.setDate(Math.min(expectedDay, lastDayOfMonth));
      safetyCounter++;
    }
  }

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

    let safetyCounter = 0;
    const MAX_ITERATIONS = 500;

    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      let specificDate: Date | null = null;

      if (frequency?.day === "TRUE") {
        specificDate = new Date(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekday === "TRUE") {
        specificDate = getNthWeekdayOfMonth(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          [1, 2, 3, 4, 5],
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekend_day === "TRUE") {
        specificDate = getNthWeekdayOfMonth(
          targetDate.getFullYear(),
          targetDate.getMonth(),
          [0, 6],
          weekMap[weekdayOfMonth]
        );
      } else {
        const keys = Object.keys(frequency);
        const dayIndices: number[] = [];
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

      if (specificDate &&
        specificDate.getTime() >= eventStartDate.getTime() &&
        specificDate.getTime() <= windowEndDate.getTime()) {
        const event = eventDataForBinding(eventDetails, specificDate);
        allEvents.push(event);
        dates.push(new Date(specificDate));
      }

      targetDate.setMonth(targetDate.getMonth() + monthFreq);
      safetyCounter++;
    }
  }

  function getNthWeekdayOfMonth(
    year: number,
    month: number,
    dayIndices: number[],
    nth: number
  ): Date | null {
    if (nth === 5) return getLastWeekdayOfMonth(year, month, dayIndices);

    const firstDayOfMonth = new Date(year, month, 1);
    const occurrences: Date[] = [];
    for (const dayIndex of dayIndices) {
      let daysToAdd = (dayIndex - firstDayOfMonth.getDay() + 7) % 7;
      if (daysToAdd === 0 && dayIndex !== firstDayOfMonth.getDay()) {
        daysToAdd = 7;
      }
      occurrences.push(new Date(year, month, 1 + daysToAdd));
    }
    occurrences.sort((a, b) => a.getTime() - b.getTime());

    if (occurrences.length === 0) return null;

    let count = 1;
    let lastDate = occurrences[0];

    let safetyCounter = 0;
    const MAX_ITERATIONS = 100;

    while (count < nth && safetyCounter < MAX_ITERATIONS) {
      const nextOccurrences: Date[] = [];
      for (const date of occurrences) {
        const nextDate: any = new Date(date);
        nextDate.setDate(nextDate.getDate() + 7);
        if (nextDate.getMonth() === month) nextOccurrences.push(nextDate);
      }
      if (nextOccurrences.length === 0) break;

      occurrences.length = 0;
      occurrences.push(...nextOccurrences);
      occurrences.sort((a, b) => a.getTime() - b.getTime());
      lastDate = occurrences[0];
      count++;
      safetyCounter++;
    }

    if (safetyCounter >= MAX_ITERATIONS) return null;
    return lastDate;
  }

  function getLastWeekdayOfMonth(
    year: number,
    month: number,
    dayIndices: number[]
  ): Date {
    const lastDayOfMonth = new Date(year, month + 1, 0);
    const daysInMonth = lastDayOfMonth.getDate();
    const lastOccurrences: Date[] = [];

    for (const dayIndex of dayIndices) {
      let day = daysInMonth;
      while (day > 0) {
        const date = new Date(year, month, day);
        if (date.getDay() === dayIndex) {
          lastOccurrences.push(date);
          break;
        }
        day--;
      }
    }
    lastOccurrences.sort((a, b) => b.getTime() - a.getTime());
    return lastOccurrences[0];
  }

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
    let targetDate = new Date(currentDate);
    targetDate.setMonth(parseInt(month) - 1);
    targetDate.setDate(parseInt(day));

    let safetyCounter = 0;
    const MAX_ITERATIONS = 200;

    while (targetDate.getTime() <= windowEndDate.getTime() && safetyCounter < MAX_ITERATIONS) {
      if (targetDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, targetDate);
        allEvents.push(event);
        dates.push(new Date(targetDate));
      }
      targetDate.setFullYear(targetDate.getFullYear() + yearFreq);
      safetyCounter++;
    }
  }

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
    const targetMonth = parseInt(month) - 1;

    let safetyCounter = 0;
    const MAX_ITERATIONS = 200;

    while (safetyCounter < MAX_ITERATIONS) {
      let specificDate: Date | null = null;

      if (frequency?.day === "TRUE") {
        specificDate = new Date(
          targetYear,
          targetMonth,
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekday === "TRUE") {
        specificDate = getNthWeekdayOfMonth(
          targetYear,
          targetMonth,
          [1, 2, 3, 4, 5],
          weekMap[weekdayOfMonth]
        );
      } else if (frequency?.weekend_day === "TRUE") {
        specificDate = getNthWeekdayOfMonth(
          targetYear,
          targetMonth,
          [0, 6],
          weekMap[weekdayOfMonth]
        );
      } else {
        const keys = Object.keys(frequency);
        const dayIndices: number[] = [];
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

      if (!specificDate || specificDate.getTime() > windowEndDate.getTime()) break;

      if (specificDate.getTime() >= eventStartDate.getTime()) {
        const event = eventDataForBinding(eventDetails, specificDate);
        allEvents.push(event);
        dates.push(new Date(specificDate));
      }

      targetYear += yearFreq;
      safetyCounter++;
    }
  }

  function getCategoryColors(category: string, isTrainingLocal: boolean) {
    if (isTrainingLocal) {
      return {
        Color: categoryOptionsColorTraining[category] ?? '#605e5c',
        FontColor: categoryOptionsFontColorTraining[category] ?? '#ffffff',
      };
    } else {
      return {
        Color: categoryOptionsColor[category] ?? '#605e5c',
        FontColor: categoryOptionsFontColor[category] ?? '#ffffff',
      };
    }
  }

  function eventDataForBinding(eventDetails: any, currentDate: Date) {
    const isTrainingLocal = calendarTitle === 'Company Calendar';
    const raw = (eventDetails?.Category && String(eventDetails.Category).trim()) ? String(eventDetails.Category) : NA_CATEGORY;
    const normCat = raw;
    const { Color, FontColor } = getCategoryColors(normCat, isTrainingLocal);

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
      RecurrenceData: eventDetails.RecurrenceData,
      Color,
      FontColor,
      category: normCat,
    };
  }

  const canEditEvent = (event: ICalendarEvent): boolean => {
    if (!userPermissions.canEdit) return false;
    return userPermissions.canEdit && (
      !event.createdBy ||
      event.createdBy.Id === currentUserId ||
      userPermissions.canDelete
    );
  };

  const canDeleteEvent = (event: ICalendarEvent): boolean => {
    if (!userPermissions.canDelete) return false;
    return userPermissions.canDelete && (
      !event.createdBy ||
      event.createdBy.Id === currentUserId ||
      userPermissions.canDelete
    );
  };

  const handleSelectEvent = (event: ICalendarEvent) => {
    if (!userPermissions.canView) return;
    setSelectedEvent(event);
    setIsNewEvent(false);
    setShowModal(true);
  };

  const handleSelectSlot = ({ start, end }: { start: Date; end: Date }) => {
    if (!userPermissions.canAdd) return;
    const newEvent: ICalendarEvent = {
      id: '',
      title: '',
      locations: '',
      startTime: start,
      endTime: end,
      description: '',
      attendees: [],
      category: NA_CATEGORY, // default to N/A to keep consistency
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

  const saveEvent = (event: ICalendarEvent) => {
    if (isNewEvent && !userPermissions.canAdd) return;
    if (!isNewEvent && !canEditEvent(event)) return;
    if (isNewEvent) createEvent(event);
    else updateEvent(event);
    setShowModal(false);
  };

  const createEvent = async (event: ICalendarEvent) => {
    if (!userPermissions.canAdd) return;
    showLoader('Loading...');
    // add your create logic...
    loadEvents();
  };

  const updateEvent = async (event: ICalendarEvent) => {
    if (!canEditEvent(event)) return;
    showLoader('Loading...');
    // add your update logic...
    loadEvents();
  };

  const deleteEvent = async (eventId: string) => {
    const event = events.find(e => e.id === eventId);
    if (!event || !canDeleteEvent(event)) return;
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
    const eventMonth = eventDate.getMonth() + 1;
    const eventYear = eventDate.getFullYear();
    return { year: eventYear, month: eventMonth };
  }

  const handleNavigate = (newDate: any, newView: any) => {
    setCurrentCalendarDate(newDate);
    setCurrentCalendarView(newView);

    if (newView === 'month') {
      const start = new Date(newDate.getFullYear(), newDate.getMonth(), 1);
      start.setDate(1 - (start.getDay() === 0 ? 7 : start.getDay()));

      const end = new Date(newDate.getFullYear(), newDate.getMonth() + 1, 0);
      const daysToAdd = 7 - end.getDay();
      end.setDate(end.getDate() + (daysToAdd === 7 ? 0 : daysToAdd));

      const filtered = AllGeneratedevents.filter((event: ICalendarEvent) => {
        return event.startTime <= end && event.endTime >= start;
      });
      hideLoader();
      setEvents(filtered);
    } else {
      hideLoader();
      setEvents(AllGeneratedevents);
    }
  };

  const handleShowMoreEvents = (eventsForDate: ICalendarEvent[], date: Date) => {
    setSelectedDateEvents(eventsForDate);
    setSelectedDate(date);
    setShowDatePanel(true);
  };

  const panelColumns: IColumn[] = React.useMemo(() => [
    {
      key: 'time',
      name: 'Time',
      minWidth: 120,
      onRender: (item: ICalendarEvent) => {
        if (item.fAllDayEvent) return 'All day';
        const start = format(new Date(item.startTime), 'h:mm a');
        const end = format(new Date(item.endTime), 'h:mm a');
        return `${start} - ${end}`;
      }
    },
    {
      key: 'title',
      name: 'Title',
      minWidth: 200,
      fieldName: 'title',
      isMultiline: true
    },
    {
      key: 'location',
      name: 'Location',
      minWidth: 180,
      onRender: (item: ICalendarEvent) => item.locations || ''
    }
  ], []);

  const components = {
    month: {
      event: (eventProps: any) => {
        const { event } = eventProps;
        return (
          <div
            className="rbc-event-preview"
            onClick={(e) => {
              e.stopPropagation();
              if (userPermissions.canView) {
                handleSelectEvent(event);
              }
            }}
            style={{ cursor: userPermissions.canView ? 'pointer' : 'default' }}
          >
            <div className="event-title">
              <span>{event.title}</span>
            </div>
          </div>
        );
      },
    },
    eventWrapper: ({ children }: any) => <>{children}</>,
  };

  const eventStyleGetter = (event: any) => {
    const style = {
      backgroundColor: event.Color,
      borderRadius: "0px",
      opacity: 0.8,
      color: event.FontColor,
      border: "0px",
      display: "block",
    };
    return { style };
  };

  if (!userPermissions.canView) {
    return (
      <div className={'modernCalendar'}>
        <h2>{calendarTitle}</h2>
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <p>You don't have permission to view this calendar.</p>
        </div>
      </div>
    );
  }

  return (
    <div className={'modernCalendar'}>
      <h2>{calendarTitle}</h2>

      {/* CATEGORY FILTER BAR */}
      <div style={{ margin: '12px 0', display: 'flex', gap: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        {/* All toggle */}
        <Checkbox
          label="All"
          checked={selectedCategories.size === categoryList.length}
          onChange={(_, checked) => {
            if (checked) setSelectedCategories(new Set(categoryList));
            else setSelectedCategories(new Set());
          }}
        />
        {/* Individual category checkboxes */}
        {categoryList.map(cat => {
          const palette = categoryPalette[cat] || { bg: '#3174ad', fg: '#fff' };
          const isChecked = selectedCategories.has(cat);
          return (
            <Checkbox
              key={cat}
              label={cat}
              checked={isChecked}
              onChange={(_, checked) => {
                setSelectedCategories(prev => {
                  const next = new Set(prev);
                  if (checked) next.add(cat);
                  else next.delete(cat);
                  return next;
                });
              }}
              styles={{
                root: { alignItems: 'center' },
                checkbox: {
                  width: 18,
                  height: 18,
                  borderRadius: 4,
                  borderColor: palette.bg,
                  background: isChecked ? palette.bg : 'transparent',
                },
                checkmark: { color: palette.fg },
                text: { marginInlineStart: 8 }
              }}
            />
          );
        })}
      </div>

      <div className={'calendarContainer'}>
        <Calendar
          localizer={localizer}
          events={filteredEvents.map(e => ({ ...e, start: e.startTime, end: e.endTime }))}
          eventPropGetter={eventStyleGetter}
          popup={false}
          onShowMore={(dayEvents, date) => {
            if (userPermissions.canView) {
              handleShowMoreEvents(dayEvents as ICalendarEvent[], date);
            }
          }}
          drilldownView={null as any}
          onDrillDown={() => {}}
          startAccessor="startTime"
          endAccessor="endTime"
          titleAccessor="title"
          style={{ height: 600 }}
          selectable={userPermissions.canAdd}
          onSelectEvent={handleSelectEvent}
          onSelectSlot={userPermissions.canAdd ? handleSelectSlot : undefined}
          onNavigate={handleNavigate}
          views={['month', 'week', 'day', 'agenda']}
          components={components}
        />
      </div>

      {showModal && (
        <EventForm
          event={selectedEvent}
          isNew={isNewEvent}
          onSave={saveEvent}
          onDelete={deleteEvent}
          Context={props?.Context}
          CalendarTitle={calendarTitle}
          MarketingCalendarId={props?.MarketingCalendarId}
          onCancel={() => setShowModal(false)}
          userPermissions={userPermissions}
          canEditEvent={selectedEvent ? canEditEvent(selectedEvent) : false}
          canDeleteEvent={selectedEvent ? canDeleteEvent(selectedEvent) : false}
          siteUrl={props?.siteUrl}
        />
      )}

      <Panel
        isOpen={showDatePanel}
        onDismiss={() => {
          setShowDatePanel(false);
          setSelectedDate(null);
          setSelectedDateEvents([]);
        }}
        headerText={selectedDate ? `Events for ${format(selectedDate, 'MMMM d, yyyy')}` : 'Events'}
        type={PanelType.medium}
      >
        <DetailsList
          items={selectedDateEvents}
          columns={panelColumns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          setKey="dateEvents"
          compact={true}
          onItemInvoked={(item) => {
            if (userPermissions.canView) {
              setShowDatePanel(false);
              handleSelectEvent(item as ICalendarEvent);
            }
          }}
        />
      </Panel>
    </div>
  );
}
