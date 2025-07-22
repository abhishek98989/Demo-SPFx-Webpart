export const parseRecurrenceToString = (recurrenceData?: string): string => {
  if (!recurrenceData) return 'No recurrence';
  
  // Handle plain text patterns first - normalize them
  if (!recurrenceData.includes('<recurrence>')) {
      return normalizeNaturalLanguagePattern(recurrenceData);
  }
  
  try {
      // Parse the recurrence data with robust error handling
      const settings = parseRecurrenceDataToSettings(recurrenceData);
      if (!settings) return 'Unrecognized recurrence pattern';
      
      let result = buildPatternDescription(settings);
      
      // Add end condition information
      result += buildEndConditionDescription(settings.rangeSettings);
      
      return result;
  } catch (error) {
      console.warn('Error parsing recurrence data:', error);
      return 'Invalid recurrence pattern';
  }
};

// Build the main pattern description
const buildPatternDescription = (settings: IRecurrenceSettings): string => {
  switch (settings.recurrenceType) {
      case 'daily':
          return buildDailyDescription(settings.pattern as IDailyRecurrencePattern);
      case 'weekly':
          return buildWeeklyDescription(settings.pattern as IWeeklyRecurrencePattern);
      case 'monthly':
          return buildMonthlyDescription(settings.pattern as IMonthlyRecurrencePattern);
      case 'yearly':
          return buildYearlyDescription(settings.pattern as IYearlyRecurrencePattern);
      default:
          return 'Custom recurrence pattern';
  }
};

// Build daily pattern description
const buildDailyDescription = (pattern: IDailyRecurrencePattern): string => {
  if (pattern.isEveryWeekday) {
      return 'Every Weekday';
  } else if (pattern.dayFrequency === 1) {
      return 'Every Day';
  } else {
      return `Every ${pattern.dayFrequency} Days`;
  }
};

// Build weekly pattern description
const buildWeeklyDescription = (pattern: IWeeklyRecurrencePattern): string => {
  const dayNames = {
      'su': 'Sun', 'mo': 'Mon', 'tu': 'Tue', 'we': 'Wed',
      'th': 'Thu', 'fr': 'Fri', 'sa': 'Sat'
  };
  
  // Handle case where no days are specified
  if (!pattern.daysOfWeek || pattern.daysOfWeek.length === 0) {
      return 'Every Week (no days specified)';
  }
  
  const daysString = pattern.daysOfWeek
      .map(day => dayNames[day.toLowerCase() as keyof typeof dayNames] || day.toUpperCase())
      .join(', ');
  
  if (pattern.weekFrequency === 1) {
      return `Every Week on ${daysString}`;
  } else {
      return `Every ${pattern.weekFrequency} Weeks on ${daysString}`;
  }
};

// Build monthly pattern description
const buildMonthlyDescription = (pattern: IMonthlyRecurrencePattern): string => {
  if (pattern.isDayOfMonth) {
      const dayWithSuffix = getDayWithSuffix(pattern.day || 1);
      if (pattern.monthFrequency === 1) {
          return `Every Month on ${dayWithSuffix}`;
      } else {
          return `Every ${pattern.monthFrequency} Months on ${dayWithSuffix}`;
      }
  } else {
      // Monthly by day pattern (e.g., first Monday, last Friday)
      const instance = getInstanceName(pattern.weekdayInstance);
      const weekday = getWeekdayName(pattern.weekdayOfMonth);
      
      if (pattern.monthFrequency === 1) {
          return `Every Month on ${instance} ${weekday}`;
      } else {
          return `Every ${pattern.monthFrequency} Months on ${instance} ${weekday}`;
      }
  }
};

// Build yearly pattern description (complete implementation)
const buildYearlyDescription = (pattern: IYearlyRecurrencePattern): string => {
  const monthNames = [
      '', 'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const monthName = monthNames[pattern.month] || `Month ${pattern.month}`;
  
  if (pattern.isSpecificDay) {
      // Yearly on specific date (e.g., "Every Year on 6th September")
      const dayWithSuffix = getDayWithSuffix(pattern.day || 1);
      if (pattern.yearFrequency === 1) {
          return `Every Year on ${dayWithSuffix} ${monthName}`;
      } else {
          return `Every ${pattern.yearFrequency} Years on ${dayWithSuffix} ${monthName}`;
      }
  } else {
      // Yearly by day pattern (e.g., "Every Year on First Monday of May")
      const instance = getInstanceName(pattern.weekdayInstance);
      const weekday = getWeekdayName(pattern.weekdayOfMonth);
      
      if (pattern.yearFrequency === 1) {
          return `Every Year on ${instance} ${weekday} of ${monthName}`;
      } else {
          return `Every ${pattern.yearFrequency} Years on ${instance} ${weekday} of ${monthName}`;
      }
  }
};

// Build end condition description
const buildEndConditionDescription = (rangeSettings: IRecurrenceRangeSettings): string => {
  if (rangeSettings.noEndDate) {
      return '';
  }
  
  if (rangeSettings.endDate) {
      const endDateStr = formatDate(rangeSettings.endDate);
      return ` until ${endDateStr}`;
  }
  
  if (rangeSettings.numOccurrences) {
      const occurrenceText = rangeSettings.numOccurrences === 1 ? 'occurrence' : 'occurrences';
      return ` for ${rangeSettings.numOccurrences} ${occurrenceText}`;
  }
  
  return '';
};

// Normalize natural language patterns
const normalizeNaturalLanguagePattern = (pattern: string): string => {
  // Basic normalization - capitalize first letter and ensure proper spacing
  return pattern.trim().replace(/\s+/g, ' ').replace(/^./, str => str.toUpperCase());
};

// Helper function to get instance name with fallbacks
const getInstanceName = (instance: number | null): string => {
  const instanceNames = ['', 'First', 'Second', 'Third', 'Fourth', 'Last'];
  
  if (instance === null || instance === undefined) {
      return 'First';
  }
  
  if (instance === 5 || instance === -1) {
      return 'Last';
  }
  
  return instanceNames[instance] || `${instance}th`;
};

// Helper function to get weekday name with fallbacks
const getWeekdayName = (weekday: number | null): string => {
  const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  if (weekday === null || weekday === undefined) {
      return 'Day';
  }
  
  return weekdayNames[weekday] || `Day ${weekday}`;
};

// Helper function to add ordinal suffix to day numbers
const getDayWithSuffix = (day: number): string => {
  if (day < 1 || day > 31) {
      return `${day}th`; // Fallback for invalid days
  }
  
  const suffix = ['th', 'st', 'nd', 'rd'][((day % 100) - 20) % 10] || 
                ['th', 'st', 'nd', 'rd'][day % 10] || 'th';
  return `${day}${suffix}`;
};

// Helper function to format date
const formatDate = (date: Date): string => {
  if (!date || isNaN(date.getTime())) {
      return 'Invalid Date';
  }
  
  const monthNames = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const day = date.getDate();
  const month = monthNames[date.getMonth()];
  const year = date.getFullYear();
  
  return `${day} ${month} ${year}`;
};

// Robust function to parse recurrence data with comprehensive error handling
const parseRecurrenceDataToSettings = (recurrenceData: string): IRecurrenceSettings | null => {
  if (!recurrenceData || !recurrenceData.includes('<recurrence>')) {
      return null;
  }

  try {
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

      // Determine recurrence type based on XML content within <repeat> tag
      if (recurrenceData.includes('<daily') || recurrenceData.includes('repeatDaily')) {
          settings.recurrenceType = 'daily';
          settings.pattern = parseDailyPattern(recurrenceData);
      } else if (recurrenceData.includes('<weekly') || recurrenceData.includes('repeatWeekly')) {
          settings.recurrenceType = 'weekly';
          settings.pattern = parseWeeklyPattern(recurrenceData);
      } else if (recurrenceData.includes('<monthly') || recurrenceData.includes('<monthlyByDay') || 
                 recurrenceData.includes('repeatMonthly') || recurrenceData.includes('repeatMonthlyByDay')) {
          settings.recurrenceType = 'monthly';
          settings.pattern = parseMonthlyPattern(recurrenceData);
      } else if (recurrenceData.includes('<yearly') || recurrenceData.includes('<yearlyByDay') ||
                 recurrenceData.includes('repeatYearly') || recurrenceData.includes('repeatYearlyByDay')) {
          settings.recurrenceType = 'yearly';
          settings.pattern = parseYearlyPattern(recurrenceData);
      } else {
          return null;
      }

      // Parse range settings with robust error handling
      settings.rangeSettings = parseRangeSettings(recurrenceData);

      return settings;
  } catch (error) {
      console.warn('Error in parseRecurrenceDataToSettings:', error);
      return null;
  }
};

// Updated parsing functions with better error handling and normalization
const parseDailyPattern = (recurrenceData: string): IDailyRecurrencePattern => {
  const pattern: IDailyRecurrencePattern = {
      isEveryWeekday: false,
      dayFrequency: 1
  };

  try {
      // Check for weekday pattern (case-insensitive)
      if (recurrenceData.toLowerCase().includes('weekday>true</weekday>') || 
          recurrenceData.includes('weekday="TRUE"') || 
          recurrenceData.includes('weekday="true"')) {
          pattern.isEveryWeekday = true;
      } else {
          // Extract day frequency from various possible formats
          const dayFreqMatch = recurrenceData.match(/<interval>(\d+)<\/interval>/) ||
                             recurrenceData.match(/dayFrequency="(\d+)"/);
          if (dayFreqMatch && dayFreqMatch[1]) {
              const frequency = parseInt(dayFreqMatch[1], 10);
              pattern.dayFrequency = frequency > 0 ? frequency : 1;
          }
      }
  } catch (error) {
      console.warn('Error parsing daily pattern:', error);
  }

  return pattern;
};

const parseWeeklyPattern = (recurrenceData: string): IWeeklyRecurrencePattern => {
  const pattern: IWeeklyRecurrencePattern = {
      weekFrequency: 1,
      daysOfWeek: []
  };

  try {
      // Extract week frequency from various possible formats
      const weekFreqMatch = recurrenceData.match(/<interval>(\d+)<\/interval>/) ||
                          recurrenceData.match(/weekFrequency="(\d+)"/);
      if (weekFreqMatch && weekFreqMatch[1]) {
          const frequency = parseInt(weekFreqMatch[1], 10);
          pattern.weekFrequency = frequency > 0 ? frequency : 1;
      }

      // Extract days of week (case-insensitive)
      const days = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
      days.forEach(day => {
          const dayPattern = new RegExp(`<${day}>TRUE<\\/${day}>|${day}="TRUE"`, 'i');
          if (dayPattern.test(recurrenceData)) {
              pattern.daysOfWeek.push(day);
          }
      });
  } catch (error) {
      console.warn('Error parsing weekly pattern:', error);
  }

  return pattern;
};

const parseMonthlyPattern = (recurrenceData: string): IMonthlyRecurrencePattern => {
  const pattern: IMonthlyRecurrencePattern = {
      monthFrequency: 1,
      day: 1,
      isDayOfMonth: true,
      weekdayOfMonth: null,
      weekdayInstance: null
  };

  try {
      // Extract month frequency from various possible formats
      const monthFreqMatch = recurrenceData.match(/monthFrequency="(\d+)"/) ||
                           recurrenceData.match(/<interval>(\d+)<\/interval>/);
      if (monthFreqMatch && monthFreqMatch[1]) {
          const frequency = parseInt(monthFreqMatch[1], 10);
          pattern.monthFrequency = frequency > 0 ? frequency : 1;
      }

      // Check if it's a monthly by day pattern
      if (recurrenceData.includes('monthlyByDay') || recurrenceData.includes('repeatMonthlyByDay')) {
          pattern.isDayOfMonth = false;
          
          // Extract weekday instance with normalization
          const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\w+)"/);
          if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
              const instanceMap: { [key: string]: number } = {
                  'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5, 'last': 5,
                  '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '-1': 5
              };
              const instanceKey = weekdayInstanceMatch[1].toLowerCase();
              pattern.weekdayInstance = instanceMap[instanceKey] || 1;
          }
          
          // Extract day of week from attributes (case-insensitive)
          const days = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
          for (let i = 0; i < days.length; i++) {
              if (recurrenceData.includes(`${days[i]}="TRUE"`) || 
                  recurrenceData.includes(`${days[i].toUpperCase()}="TRUE"`)) {
                  pattern.weekdayOfMonth = i;
                  break;
              }
          }
      } else {
          // Extract day of month from various formats
          const dayMatch = recurrenceData.match(/day="(\d+)"/) ||
                         recurrenceData.match(/<day>(\d+)<\/day>/);
          if (dayMatch && dayMatch[1]) {
              const day = parseInt(dayMatch[1], 10);
              pattern.day = (day >= 1 && day <= 31) ? day : 1;
          }
      }
  } catch (error) {
      console.warn('Error parsing monthly pattern:', error);
  }

  return pattern;
};

const parseYearlyPattern = (recurrenceData: string): IYearlyRecurrencePattern => {
  const pattern: IYearlyRecurrencePattern = {
      yearFrequency: 1,
      month: 1,
      day: 1,
      isSpecificDay: true,
      weekdayOfMonth: null,
      weekdayInstance: null
  };

  try {
      // Extract year frequency from various possible formats
      const yearFreqMatch = recurrenceData.match(/yearFrequency="(\d+)"/) ||
                          recurrenceData.match(/<interval>(\d+)<\/interval>/);
      if (yearFreqMatch && yearFreqMatch[1]) {
          const frequency = parseInt(yearFreqMatch[1], 10);
          pattern.yearFrequency = frequency > 0 ? frequency : 1;
      }

      // Extract month
      const monthMatch = recurrenceData.match(/month="(\d+)"/) ||
                       recurrenceData.match(/<month>(\d+)<\/month>/);
      if (monthMatch && monthMatch[1]) {
          const month = parseInt(monthMatch[1], 10);
          pattern.month = (month >= 1 && month <= 12) ? month : 1;
      }

      // Check if it's a yearly by day pattern
      if (recurrenceData.includes('yearlyByDay') || 
          recurrenceData.includes('repeatYearlyByDay') ||
          recurrenceData.includes('weekdayOfMonth')) {
          pattern.isSpecificDay = false;
          
          // Extract weekday instance
          const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\w+)"/) ||
                                     recurrenceData.match(/<weekdayOfMonth>(\d+)<\/weekdayOfMonth>/);
          if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
              const instanceMap: { [key: string]: number } = {
                  'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5, 'last': 5,
                  '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '-1': 5
              };
              const instanceKey = weekdayInstanceMatch[1].toLowerCase();
              pattern.weekdayInstance = instanceMap[instanceKey] || 
                                      parseInt(weekdayInstanceMatch[1], 10) || 1;
          }
          
          // Extract day of week
          const weekdayMatch = recurrenceData.match(/<day>(\d+)<\/day>/) ||
                             recurrenceData.match(/day="(\d+)"/);
          if (weekdayMatch && weekdayMatch[1]) {
              const weekday = parseInt(weekdayMatch[1], 10);
              pattern.weekdayOfMonth = (weekday >= 0 && weekday <= 6) ? weekday : 0;
          }
      } else {
          // Extract day of month
          const dayMatch = recurrenceData.match(/day="(\d+)"/) ||
                         recurrenceData.match(/<day>(\d+)<\/day>/);
          if (dayMatch && dayMatch[1]) {
              const day = parseInt(dayMatch[1], 10);
              pattern.day = (day >= 1 && day <= 31) ? day : 1;
          }
      }
  } catch (error) {
      console.warn('Error parsing yearly pattern:', error);
  }

  return pattern;
};

const parseRangeSettings = (recurrenceData: string): IRecurrenceRangeSettings => {
  const settings: IRecurrenceRangeSettings = {
      startDate: new Date(),
      endDate: null,
      numOccurrences: null,
      noEndDate: true
  };

  try {
      // Extract start date from windowStart element
      const startDateMatch = recurrenceData.match(/<windowStart>([^<]+)<\/windowStart>/);
      if (startDateMatch && startDateMatch[1]) {
          const startDate = new Date(startDateMatch[1]);
          if (!isNaN(startDate.getTime())) {
              settings.startDate = startDate;
          }
      }

      // Check for repeat forever flag first
      const repeatForeverMatch = recurrenceData.match(/<repeatForever>(TRUE|FALSE)<\/repeatForever>/i);
      if (repeatForeverMatch) {
          if (repeatForeverMatch[1].toUpperCase() === 'TRUE') {
              settings.noEndDate = true;
              settings.endDate = null;
              settings.numOccurrences = null;
              return settings; // Exit early if it repeats forever
          } else {
              settings.noEndDate = false;
          }
      }

      // Check for end date in windowEnd element
      const endDateMatch = recurrenceData.match(/<windowEnd>([^<]+)<\/windowEnd>/);
      if (endDateMatch && endDateMatch[1]) {
          const endDate = new Date(endDateMatch[1]);
          if (!isNaN(endDate.getTime())) {
              settings.noEndDate = false;
              settings.endDate = endDate;
          }
      }

      // Check for number of occurrences in repeatInstances element
      const numOccurrencesMatch = recurrenceData.match(/<repeatInstances>(\d+)<\/repeatInstances>/);
      if (numOccurrencesMatch && numOccurrencesMatch[1]) {
          const numOccurrences = parseInt(numOccurrencesMatch[1], 10);
          if (numOccurrences > 0) {
              settings.noEndDate = false;
              settings.numOccurrences = numOccurrences;
              // Clear end date if we have occurrence count
              settings.endDate = null;
          }
      }

      // If no end conditions found and repeatForever is FALSE, default to no end date
      if (settings.noEndDate === false && !settings.endDate && !settings.numOccurrences) {
          settings.noEndDate = true;
      }
  } catch (error) {
      console.warn('Error parsing range settings:', error);
  }

  return settings;
};

// Type definitions (unchanged)
interface IRecurrenceSettings {
recurrenceType: string;
pattern: IDailyRecurrencePattern | IWeeklyRecurrencePattern | IMonthlyRecurrencePattern | IYearlyRecurrencePattern;
rangeSettings: IRecurrenceRangeSettings;
}

interface IDailyRecurrencePattern {
isEveryWeekday: boolean;
dayFrequency: number;
}

interface IWeeklyRecurrencePattern {
weekFrequency: number;
daysOfWeek: string[];
}

interface IMonthlyRecurrencePattern {
monthFrequency: number;
day?: number;
isDayOfMonth: boolean;
weekdayOfMonth: number | null;
weekdayInstance: number | null;
}

interface IYearlyRecurrencePattern {
yearFrequency: number;
month: number;
day?: number;
isSpecificDay: boolean;
weekdayOfMonth: number | null;
weekdayInstance: number | null;
}

interface IRecurrenceRangeSettings {
startDate: Date;
endDate: Date | null;
numOccurrences: number | null;
noEndDate: boolean;
}