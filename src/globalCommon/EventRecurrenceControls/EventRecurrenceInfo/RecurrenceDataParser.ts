/**
 * Utility class to parse SharePoint recurrence XML data
 */
export class RecurrenceDataParser {
  /**
   * Parse recurrence data XML string and return structured recurrence settings
   * @param recurrenceData XML string containing recurrence pattern
   * @returns Parsed recurrence settings object
   */
  public static parseRecurrenceData(recurrenceData: string): IRecurrenceSettings | null {
    if (!recurrenceData) return null;

    const settings: IRecurrenceSettings = {
      recurrenceType: '',
      pattern: {
        isEveryWeekday: false,
        dayFrequency: 1
      },
      rangeSettings: {
        startDate: new Date(),
        endDate: null,
        numOccurrences: null,
        noEndDate: true
      }
    };

    // Extract recurrence type
    if (recurrenceData.indexOf('<daily') !== -1) {
      settings.recurrenceType = 'daily';
      settings.pattern = this.parseDailyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<weekly') !== -1) {
      settings.recurrenceType = 'weekly';
      settings.pattern = this.parseWeeklyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<monthly') !== -1 || recurrenceData.indexOf('<monthlyByDay') !== -1) {
      settings.recurrenceType = 'monthly';
      settings.pattern = this.parseMonthlyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<yearly') !== -1) {
      settings.recurrenceType = 'yearly';
      settings.pattern = this.parseYearlyPattern(recurrenceData);
    } else {
      return null; // Unknown recurrence type
    }

    // Parse range settings
    settings.rangeSettings = this.parseRangeSettings(recurrenceData);

    return settings;
  }

  /**
   * Parse daily recurrence pattern
   */
  private static parseDailyPattern(recurrenceData: string): IDailyRecurrencePattern {
    const pattern: IDailyRecurrencePattern = {
      isEveryWeekday: false,
      dayFrequency: 1
    };

    // Check if it's every weekday
    if (recurrenceData.indexOf('weekday="TRUE"') !== -1) {
      pattern.isEveryWeekday = true;
    } else {
      // Extract day frequency
      const dayFreqMatch = recurrenceData.match(/dayFrequency="(\d+)"/);
      if (dayFreqMatch && dayFreqMatch[1]) {
        pattern.dayFrequency = parseInt(dayFreqMatch[1], 10);
      }
    }

    return pattern;
  }

  /**
   * Parse weekly recurrence pattern
   */
  private static parseWeeklyPattern(recurrenceData: string): IWeeklyRecurrencePattern {
    const pattern: IWeeklyRecurrencePattern = {
      weekFrequency: 1,
      daysOfWeek: []
    };

    // Extract week frequency
    const weekFreqMatch = recurrenceData.match(/weekFrequency="(\d+)"/);
    if (weekFreqMatch && weekFreqMatch[1]) {
      pattern.weekFrequency = parseInt(weekFreqMatch[1], 10);
    }

    // Extract days of week
    const days = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
    days.forEach(day => {
      if (recurrenceData.indexOf(`${day}="TRUE"`) !== -1) {
        pattern.daysOfWeek.push(day);
      }
    });

    return pattern;
  }

  /**
   * Parse monthly recurrence pattern
   */
  private static parseMonthlyPattern(recurrenceData: string): IMonthlyRecurrencePattern {
    const pattern: IMonthlyRecurrencePattern = {
      monthFrequency: 1,
      day: 1,
      isDayOfMonth: true,
      weekdayOfMonth: null,
      weekdayInstance: null
    };

    // Check if it's monthly by day pattern
    if (recurrenceData.indexOf('<monthlyByDay') !== -1) {
      pattern.isDayOfMonth = false;
      
      // Extract month frequency
      const monthFreqMatch = recurrenceData.match(/monthFrequency="(\d+)"/);
      if (monthFreqMatch && monthFreqMatch[1]) {
        pattern.monthFrequency = parseInt(monthFreqMatch[1], 10);
      }
      
      // Extract weekday instance (first, second, third, fourth, last)
      const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\d+)"/);
      if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
        pattern.weekdayInstance = parseInt(weekdayInstanceMatch[1], 10);
      }
      
      // Extract weekday (0=Sunday, 1=Monday, etc.)
      const weekdayMatch = recurrenceData.match(/day="(\d+)"/);
      if (weekdayMatch && weekdayMatch[1]) {
        pattern.weekdayOfMonth = parseInt(weekdayMatch[1], 10);
      }
    } else {
      // Regular monthly pattern
      // Extract month frequency
      const monthFreqMatch = recurrenceData.match(/monthFrequency="(\d+)"/);
      if (monthFreqMatch && monthFreqMatch[1]) {
        pattern.monthFrequency = parseInt(monthFreqMatch[1], 10);
      }
      
      // Extract day of month
      const dayMatch = recurrenceData.match(/day="(\d+)"/);
      if (dayMatch && dayMatch[1]) {
        pattern.day = parseInt(dayMatch[1], 10);
      }
    }

    return pattern;
  }

  /**
   * Parse yearly recurrence pattern
   */
  private static parseYearlyPattern(recurrenceData: string): IYearlyRecurrencePattern {
    const pattern: IYearlyRecurrencePattern = {
      yearFrequency: 1,
      month: 1,
      day: 1,
      isSpecificDay: true,
      weekdayOfMonth: null,
      weekdayInstance: null
    };

    // Extract year frequency
    const yearFreqMatch = recurrenceData.match(/yearFrequency="(\d+)"/);
    if (yearFreqMatch && yearFreqMatch[1]) {
      pattern.yearFrequency = parseInt(yearFreqMatch[1], 10);
    }

    // Extract month
    const monthMatch = recurrenceData.match(/month="(\d+)"/);
    if (monthMatch && monthMatch[1]) {
      pattern.month = parseInt(monthMatch[1], 10);
    }

    // Check if it's by month day or by weekday
    if (recurrenceData.indexOf('weekdayOfMonth') !== -1) {
      pattern.isSpecificDay = false;
      
      // Extract weekday instance (first, second, third, fourth, last)
      const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\d+)"/);
      if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
        pattern.weekdayInstance = parseInt(weekdayInstanceMatch[1], 10);
      }
      
      // Extract weekday (0=Sunday, 1=Monday, etc.)
      const weekdayMatch = recurrenceData.match(/day="(\d+)"/);
      if (weekdayMatch && weekdayMatch[1]) {
        pattern.weekdayOfMonth = parseInt(weekdayMatch[1], 10);
      }
    } else {
      // Extract day of month
      const dayMatch = recurrenceData.match(/day="(\d+)"/);
      if (dayMatch && dayMatch[1]) {
        pattern.day = parseInt(dayMatch[1], 10);
      }
    }

    return pattern;
  }

  /**
   * Parse recurrence range settings
   */
  private static parseRangeSettings(recurrenceData: string): IRecurrenceRangeSettings {
    const settings: IRecurrenceRangeSettings = {
      startDate: new Date(),
      endDate: null,
      numOccurrences: null,
      noEndDate: true
    };

    // Extract start date
    const startDateMatch = recurrenceData.match(/start="([^"]+)"/);
    if (startDateMatch && startDateMatch[1]) {
      settings.startDate = new Date(startDateMatch[1]);
    }

    // Check for end type
    if (recurrenceData.indexOf('endDate="') !== -1) {
      // Has end date
      settings.noEndDate = false;
      const endDateMatch = recurrenceData.match(/endDate="([^"]+)"/);
      if (endDateMatch && endDateMatch[1]) {
        settings.endDate = new Date(endDateMatch[1]);
      }
    } else if (recurrenceData.indexOf('numOccurrences="') !== -1) {
      // Has number of occurrences
      settings.noEndDate = false;
      const numOccurrencesMatch = recurrenceData.match(/numOccurrences="(\d+)"/);
      if (numOccurrencesMatch && numOccurrencesMatch[1]) {
        settings.numOccurrences = parseInt(numOccurrencesMatch[1], 10);
      }
    }

    return settings;
  }
}

/**
 * Interface for parsed recurrence settings
 */
export interface IRecurrenceSettings {
  recurrenceType: string;
  pattern: IDailyRecurrencePattern | IWeeklyRecurrencePattern | IMonthlyRecurrencePattern | IYearlyRecurrencePattern;
  rangeSettings: IRecurrenceRangeSettings;
}

/**
 * Interface for daily recurrence pattern
 */
export interface IDailyRecurrencePattern {
  isEveryWeekday: boolean;
  dayFrequency: number;
}

/**
 * Interface for weekly recurrence pattern
 */
export interface IWeeklyRecurrencePattern {
  weekFrequency: number;
  daysOfWeek: string[];
}

/**
 * Interface for monthly recurrence pattern
 */
export interface IMonthlyRecurrencePattern {
  monthFrequency: number;
  day?: number;
  isDayOfMonth: boolean;
  weekdayOfMonth: number | null;
  weekdayInstance: number | null;
}

/**
 * Interface for yearly recurrence pattern
 */
export interface IYearlyRecurrencePattern {
  yearFrequency: number;
  month: number;
  day?: number;
  isSpecificDay: boolean;
  weekdayOfMonth: number | null;
  weekdayInstance: number | null;
}

/**
 * Interface for recurrence range settings
 */
export interface IRecurrenceRangeSettings {
  startDate: Date;
  endDate: Date | null;
  numOccurrences: number | null;
  noEndDate: boolean;
}