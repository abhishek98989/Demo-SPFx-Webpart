export const parseRecurrenceToString = (recurrenceData?: string): string => {
    if (!recurrenceData) return 'No recurrence';
    
    // Parse the recurrence data using the RecurrenceDataParser logic
    const settings = parseRecurrenceDataToSettings(recurrenceData);
    if (!settings) return 'Invalid recurrence pattern';
    
    let result = '';
    
    // Build the pattern description based on recurrence type
    switch (settings.recurrenceType) {
      case 'daily':
        const dailyPattern = settings.pattern as IDailyRecurrencePattern;
        if (dailyPattern.isEveryWeekday) {
          result = 'Every Weekday';
        } else if (dailyPattern.dayFrequency === 1) {
          result = 'Every Day';
        } else {
          result = `Every ${dailyPattern.dayFrequency} Days`;
        }
        break;
        
      case 'weekly':
        const weeklyPattern = settings.pattern as IWeeklyRecurrencePattern;
        const dayNames = {
          'su': 'Sun', 'mo': 'Mon', 'tu': 'Tue', 'we': 'Wed',
          'th': 'Thu', 'fr': 'Fri', 'sa': 'Sat'
        };
        
        const daysString = weeklyPattern.daysOfWeek
          .map(day => dayNames[day as keyof typeof dayNames])
          .join(', ');
        
        if (weeklyPattern.weekFrequency === 1) {
          result = `Every Week on ${daysString}`;
        } else {
          result = `Every ${weeklyPattern.weekFrequency} Weeks on ${daysString}`;
        }
        break;
        
      case 'monthly':
        const monthlyPattern = settings.pattern as IMonthlyRecurrencePattern;
        if (monthlyPattern.isDayOfMonth) {
          const dayWithSuffix = getDayWithSuffix(monthlyPattern.day || 1);
          if (monthlyPattern.monthFrequency === 1) {
            result = `Every Month on ${dayWithSuffix}`;
          } else {
            result = `Every ${monthlyPattern.monthFrequency} Months on ${dayWithSuffix}`;
          }
        } else {
          // Monthly by day pattern (e.g., first Monday, last Friday)
          const instanceNames = ['', 'First', 'Second', 'Third', 'Fourth', 'Last'];
          const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          
          const instance = monthlyPattern.weekdayInstance === 5 ? 'Last' : 
                          instanceNames[monthlyPattern.weekdayInstance || 1];
          const weekday = weekdayNames[monthlyPattern.weekdayOfMonth || 0];
          
          if (monthlyPattern.monthFrequency === 1) {
            result = `Every Month on ${instance} ${weekday}`;
          } else {
            result = `Every ${monthlyPattern.monthFrequency} Months on ${instance} ${weekday}`;
          }
        }
        break;
        
      case 'yearly':
        const yearlyPattern = settings.pattern as IYearlyRecurrencePattern;
        const monthNames = [
          '', 'January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December'
        ];
        
        if (yearlyPattern.isSpecificDay) {
          const dayWithSuffix = getDayWithSuffix(yearlyPattern.day || 1);
          const monthName = monthNames[yearlyPattern.month];
          if (yearlyPattern.yearFrequency === 1) {
            result = `Every Year on ${dayWithSuffix} ${monthName}`;
          } else {
            result = `Every ${yearlyPattern.yearFrequency} Years on ${dayWithSuffix} ${monthName}`;
          }
        } else {
          // Yearly by day pattern (e.g., first Monday of May)
          const instanceNames = ['', 'First', 'Second', 'Third', 'Fourth', 'Last'];
          const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          
          const instance = yearlyPattern.weekdayInstance === 5 ? 'Last' : 
                          instanceNames[yearlyPattern.weekdayInstance || 1];
          const weekday = weekdayNames[yearlyPattern.weekdayOfMonth || 0];
          const monthName = monthNames[yearlyPattern.month];
          
          if (yearlyPattern.yearFrequency === 1) {
            result = `Every Year on ${instance} ${weekday} of ${monthName}`;
          } else {
            result = `Every ${yearlyPattern.yearFrequency} Years on ${instance} ${weekday} of ${monthName}`;
          }
        }
        break;
        
      default:
        return 'Custom recurrence pattern';
    }
    
    // Add range information
    const rangeSettings = settings.rangeSettings;
    if (!rangeSettings.noEndDate) {
      if (rangeSettings.endDate) {
        const endDateStr = formatDate(rangeSettings.endDate);
        result += ` till ${endDateStr}`;
      } else if (rangeSettings.numOccurrences) {
        result += ` for ${rangeSettings.numOccurrences} occurrences`;
      }
    }
    
    return result;
  };
  
  // Helper function to add ordinal suffix to day numbers
  const getDayWithSuffix = (day: number): string => {
    const suffix = ['th', 'st', 'nd', 'rd'][((day % 100) - 20) % 10] || 
                  ['th', 'st', 'nd', 'rd'][day % 10] || 'th';
    return `${day}${suffix}`;
  };
  
  // Helper function to format date
  const formatDate = (date: Date): string => {
    const monthNames = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    const day = date.getDate();
    const month = monthNames[date.getMonth()];
    const year = date.getFullYear();
    
    return `${day} ${month} ${year}`;
  };
  
  // Helper function to parse recurrence data (simplified version of RecurrenceDataParser.parseRecurrenceData)
  const parseRecurrenceDataToSettings = (recurrenceData: string): IRecurrenceSettings | null => {
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
  
    // Extract recurrence type and parse pattern
    if (recurrenceData.indexOf('<daily') !== -1) {
      settings.recurrenceType = 'daily';
      settings.pattern = parseDailyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<weekly') !== -1) {
      settings.recurrenceType = 'weekly';
      settings.pattern = parseWeeklyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<monthly') !== -1 || recurrenceData.indexOf('<monthlyByDay') !== -1) {
      settings.recurrenceType = 'monthly';
      settings.pattern = parseMonthlyPattern(recurrenceData);
    } else if (recurrenceData.indexOf('<yearly') !== -1) {
      settings.recurrenceType = 'yearly';
      settings.pattern = parseYearlyPattern(recurrenceData);
    } else {
      return null;
    }
  
    // Parse range settings
    settings.rangeSettings = parseRangeSettings(recurrenceData);
  
    return settings;
  };
  
  // Simplified parsing functions (based on RecurrenceDataParser)
  const parseDailyPattern = (recurrenceData: string): IDailyRecurrencePattern => {
    const pattern: IDailyRecurrencePattern = {
      isEveryWeekday: false,
      dayFrequency: 1
    };
  
    if (recurrenceData.indexOf('weekday="TRUE"') !== -1) {
      pattern.isEveryWeekday = true;
    } else {
      const dayFreqMatch = recurrenceData.match(/dayFrequency="(\d+)"/);
      if (dayFreqMatch && dayFreqMatch[1]) {
        pattern.dayFrequency = parseInt(dayFreqMatch[1], 10);
      }
    }
  
    return pattern;
  };
  
  const parseWeeklyPattern = (recurrenceData: string): IWeeklyRecurrencePattern => {
    const pattern: IWeeklyRecurrencePattern = {
      weekFrequency: 1,
      daysOfWeek: []
    };
  
    const weekFreqMatch = recurrenceData.match(/weekFrequency="(\d+)"/);
    if (weekFreqMatch && weekFreqMatch[1]) {
      pattern.weekFrequency = parseInt(weekFreqMatch[1], 10);
    }
  
    const days = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
    days.forEach(day => {
      if (recurrenceData.indexOf(`${day}="TRUE"`) !== -1) {
        pattern.daysOfWeek.push(day);
      }
    });
  
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
  
    if (recurrenceData.indexOf('<monthlyByDay') !== -1) {
      pattern.isDayOfMonth = false;
      
      const monthFreqMatch = recurrenceData.match(/monthFrequency="(\d+)"/);
      if (monthFreqMatch && monthFreqMatch[1]) {
        pattern.monthFrequency = parseInt(monthFreqMatch[1], 10);
      }
      
      const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\d+)"/);
      if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
        pattern.weekdayInstance = parseInt(weekdayInstanceMatch[1], 10);
      }
      
      const weekdayMatch = recurrenceData.match(/day="(\d+)"/);
      if (weekdayMatch && weekdayMatch[1]) {
        pattern.weekdayOfMonth = parseInt(weekdayMatch[1], 10);
      }
    } else {
      const monthFreqMatch = recurrenceData.match(/monthFrequency="(\d+)"/);
      if (monthFreqMatch && monthFreqMatch[1]) {
        pattern.monthFrequency = parseInt(monthFreqMatch[1], 10);
      }
      
      const dayMatch = recurrenceData.match(/day="(\d+)"/);
      if (dayMatch && dayMatch[1]) {
        pattern.day = parseInt(dayMatch[1], 10);
      }
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
  
    const yearFreqMatch = recurrenceData.match(/yearFrequency="(\d+)"/);
    if (yearFreqMatch && yearFreqMatch[1]) {
      pattern.yearFrequency = parseInt(yearFreqMatch[1], 10);
    }
  
    const monthMatch = recurrenceData.match(/month="(\d+)"/);
    if (monthMatch && monthMatch[1]) {
      pattern.month = parseInt(monthMatch[1], 10);
    }
  
    if (recurrenceData.indexOf('weekdayOfMonth') !== -1) {
      pattern.isSpecificDay = false;
      
      const weekdayInstanceMatch = recurrenceData.match(/weekdayOfMonth="(\d+)"/);
      if (weekdayInstanceMatch && weekdayInstanceMatch[1]) {
        pattern.weekdayInstance = parseInt(weekdayInstanceMatch[1], 10);
      }
      
      const weekdayMatch = recurrenceData.match(/day="(\d+)"/);
      if (weekdayMatch && weekdayMatch[1]) {
        pattern.weekdayOfMonth = parseInt(weekdayMatch[1], 10);
      }
    } else {
      const dayMatch = recurrenceData.match(/day="(\d+)"/);
      if (dayMatch && dayMatch[1]) {
        pattern.day = parseInt(dayMatch[1], 10);
      }
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
  
    const startDateMatch = recurrenceData.match(/start="([^"]+)"/);
    if (startDateMatch && startDateMatch[1]) {
      settings.startDate = new Date(startDateMatch[1]);
    }
  
    if (recurrenceData.indexOf('endDate="') !== -1) {
      settings.noEndDate = false;
      const endDateMatch = recurrenceData.match(/endDate="([^"]+)"/);
      if (endDateMatch && endDateMatch[1]) {
        settings.endDate = new Date(endDateMatch[1]);
      }
    } else if (recurrenceData.indexOf('numOccurrences="') !== -1) {
      settings.noEndDate = false;
      const numOccurrencesMatch = recurrenceData.match(/numOccurrences="(\d+)"/);
      if (numOccurrencesMatch && numOccurrencesMatch[1]) {
        settings.numOccurrences = parseInt(numOccurrencesMatch[1], 10);
      }
    }
  
    return settings;
  };
  
  // Type definitions (same as in RecurrenceDataParser.ts)
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