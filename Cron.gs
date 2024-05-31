function testParseCronField() {
  Logger.log(PARSE_CRON_FIELD("1-5", 1, 5));  // Should log [1, 2, 3, 4, 5]
  Logger.log(PARSE_CRON_FIELD("3-/2", 1, 10));  // Should log [3, 5, 7, 9]
  Logger.log(PARSE_CRON_FIELD("1,2,5", 1, 5));  // Should log [1, 2, 5]
}

/**
 * Parses a cron field to an array of numerical values.
 * @param {string} field - The cron field.
 * @param {number} min - The minimum value for the field.
 * @param {number} max - The maximum value for the field.
 * @returns {number[]} - Array of valid numerical values for the field.
 */
function PARSE_CRON_FIELD(field, min, max) {
  let result = new Set();

  // Handle wildcard (*)
  if (field === '*') {
    for (let i = min; i <= max; i++) result.add(i);
    return Array.from(result);
  }

  // Handle lists and ranges (1,3,5 or 1-5)
  let parts = field.split(',');
  for (let part of parts) {
    
    if (part.indexOf('/') > -1) {
      // Handle ranges with steps (1-15/2)
      let [range, step] = part.split('/');
      step = parseInt(step, 10);

      if (range === '*') {
        for (let i = min; i <= max; i += step) result.add(i);
      } else if (range.indexOf('-') > -1) {
        // Range with step
        let [start, end] = range.split('-');
        start = start === '' ? min : Number(start); // Ensure start is defined
        end = end === '' ? max : Number(end); // If end is omitted, assume max
        for (let i = start; i <= end; i += step) result.add(i);
      } else {
        // Interval with step (e.g. 5/2)
        let start = parseInt(range, 10);
        for (let i = start; i <= max; i += step) result.add(i);
      }
    } else if (part.indexOf('-') > -1) {
      // Range without steps (1-5)
      let [start, end] = part.split('-').map(Number);
      end = isNaN(end) ? max : end; // If end is omitted, assume max
      start = isNaN(start) ? min : start; // Ensure start is defined
      for (let i = start; i <= end; i++) result.add(i);
    } else {
      // Individual value
      result.add(parseInt(part, 10));
    }
  }

  return Array.from(result);
}

/**
 * Checks if a given date matches the cron criteria.
 * @param {Date} date - The date to check.
 * @param {number[]} daysOfMonth - Array of valid days of the month.
 * @param {number[]} months - Array of valid months.
 * @param {number[]} daysOfWeek - Array of valid days of the week.
 * @param {number[]} daysSinceEpoch - Array of valid days since the epoch (if not a wildcard).
 * @returns {boolean} - True if the date matches the criteria, false otherwise.
 */
function matchesCronCriteria(date, daysOfMonth, months, daysOfWeek, daysSinceEpoch) {
  const dayOfMonth = date.getDate();
  const month = date.getMonth() + 1; // getMonth() is zero-based
  const dayOfWeek = date.getDay();

  if (!daysOfMonth.includes(dayOfMonth)) return false;
  if (!months.includes(month)) return false;
  if (!daysOfWeek.includes(dayOfWeek)) return false;

  if (daysSinceEpoch.length > 0) {
    const epochRange = SpreadsheetApp.getActive().getNamedRanges().find(r => r.getName() === 'EpochZero');
    if (!epochRange) throw 'Could not find named range "EpochZero"!';
    const epoch = epochRange.getRange().getValue();
    const daysSinceEpochValue = Math.floor((date - epoch) / (1000 * 60 * 60 * 24));
    if (!daysSinceEpoch.includes(daysSinceEpochValue)) return false;
  }

  return true;
}

/**
 * Custom function for Google Sheets to get the next execution date for a given cron expression.
 * @param {string} cronExpression - The cron expression with 4 fields excluding minutes and hours.
 * @param {string} referenceDate - The reference date in ISO format (yyyy-mm-dd) to start the search from.
 * @returns {string} - The next execution date in ISO format without time.
 */
function CRON_NEXT_DATE(cronExpression, referenceDate) {
  const now = new Date(referenceDate);
  now.setMilliseconds(0);
  now.setSeconds(0);
  now.setMinutes(0);
  now.setHours(0); // Start from the beginning of the reference date

  // Define the ranges for cron fields
  const ranges = {
    dayOfMonth: { min: 1, max: 31 },
    month: { min: 1, max: 12 },
    dayOfWeek: { min: 0, max: 6 }
  };

  // Split the cron expression into fields
  const [dayOfMonthField, monthField, dayOfWeekField, daysSinceEpochField] = cronExpression.split(' ');

  // Parse each field
  const daysOfMonth = PARSE_CRON_FIELD(dayOfMonthField, ranges.dayOfMonth.min, ranges.dayOfMonth.max);
  const months = PARSE_CRON_FIELD(monthField, ranges.month.min, ranges.month.max);
  const daysOfWeek = PARSE_CRON_FIELD(dayOfWeekField, ranges.dayOfWeek.min, ranges.dayOfWeek.max);

  // Prepare daysSinceEpoch array if not a wildcard
  const daysSinceEpoch = daysSinceEpochField !== '*' ? PARSE_CRON_FIELD(daysSinceEpochField, 0, 365*10) : [];

  // Search for the next valid date within one year
  let nextDate = new Date(now);
  for (let i = 0; i < 365; i++) {
    if (matchesCronCriteria(nextDate, daysOfMonth, months, daysOfWeek, daysSinceEpoch)) {
      return nextDate.toISOString().split('T')[0]; // Return only the date part without the time
    }
    nextDate.setDate(nextDate.getDate() + 1); // Move to the next day
  }

  return ''; // Fallback if no valid date is found within the next year
}

/**
 * Custom function for Google Sheets to get the latest date before the given reference date that satisfies the cron criteria.
 * @param {string} cronExpression - The cron expression with 4 fields excluding minutes and hours.
 * @param {string} referenceDate - The reference date in ISO format (yyyy-mm-dd) to start the search from.
 * @returns {string} - The latest execution date in ISO format without time.
 */
function CRON_LAST_DATE(cronExpression, referenceDate) {
  const now = new Date(referenceDate);
  now.setMilliseconds(0);
  now.setSeconds(0);
  now.setMinutes(0);
  now.setHours(0); // Start from the beginning of the reference date
  
  // Define the ranges for cron fields
  const ranges = {
    dayOfMonth: { min: 1, max: 31 },
    month: { min: 1, max: 12 },
    dayOfWeek: { min: 0, max: 6 }
  };

  // Split the cron expression into fields
  const [dayOfMonthField, monthField, dayOfWeekField, daysSinceEpochField] = cronExpression.split(' ');

  // Parse each field
  const daysOfMonth = PARSE_CRON_FIELD(dayOfMonthField, ranges.dayOfMonth.min, ranges.dayOfMonth.max);
  const months = PARSE_CRON_FIELD(monthField, ranges.month.min, ranges.month.max);
  const daysOfWeek = PARSE_CRON_FIELD(dayOfWeekField, ranges.dayOfWeek.min, ranges.dayOfWeek.max);

  // Prepare daysSinceEpoch array if not a wildcard
  const daysSinceEpoch = daysSinceEpochField !== '*' ? PARSE_CRON_FIELD(daysSinceEpochField, 0, 365*500) : [];

  // Search for the last valid date
  let lastDate = new Date(now);
  for (let i = 0; i < 365; i++) {
    lastDate.setDate(lastDate.getDate() - 1); // Move to the previous day
    if (matchesCronCriteria(lastDate, daysOfMonth, months, daysOfWeek, daysSinceEpoch)) {
      return lastDate.toISOString().split('T')[0]; // Return only the date part without the time
    }
  }

  return ''; // Fallback if no valid date is found within the past year
}

/**
 * Custom function to count the number of dates between startDate (inclusive) and endDate (exclusive) 
 * that match the given cron expression.
 * @param {string} cronExpression - The cron expression with 4 fields excluding minutes and hours.
 * @param {string} startDate - The start date in ISO format (yyyy-mm-dd).
 * @param {string} endDate - The end date in ISO format (yyyy-mm-dd).
 * @returns {number} - The count of dates between startDate and endDate that match the cron expression.
 */
function CRON_COUNT(cronExpression, startDate, endDate) {
  let start = new Date(startDate);
  let end = new Date(endDate);

  const ranges = {
    dayOfMonth: { min: 1, max: 31 },
    month: { min: 1, max: 12 },
    dayOfWeek: { min: 0, max: 6 }
  };

  const [dayOfMonthField, monthField, dayOfWeekField, daysSinceEpochField] = cronExpression.split(' ');

  const daysOfMonth = PARSE_CRON_FIELD(dayOfMonthField, ranges.dayOfMonth.min, ranges.dayOfMonth.max);
  const months = PARSE_CRON_FIELD(monthField, ranges.month.min, ranges.month.max);
  const daysOfWeek = PARSE_CRON_FIELD(dayOfWeekField, ranges.dayOfWeek.min, ranges.dayOfWeek.max);
  const daysSinceEpoch = daysSinceEpochField !== '*' ? PARSE_CRON_FIELD(daysSinceEpochField, 0, 365*500) : [];

  let count = 0;

  // Create a helper to quickly check if a given day/month combination matches the cron criteria.
  function matchesCronCriteria(date) {
    const day = date.getDate();
    const month = date.getMonth() + 1; // JavaScript months are 0-based
    const dayOfWeek = date.getDay();
    const daysSinceEpochCount = Math.floor((date - new Date(0)) / (1000 * 60 * 60 * 24));

    return (daysOfMonth.includes(day) || dayOfMonthField === '*') &&
           (months.includes(month) || monthField === '*') &&
           (daysOfWeek.includes(dayOfWeek) || dayOfWeekField === '*') &&
           (daysSinceEpoch.length === 0 || daysSinceEpoch.includes(daysSinceEpochCount));
  }

  // Iterate through each date within the range, checking only the dates within our parsed ranges
  for (let currentDate = start; currentDate < end; currentDate.setDate(currentDate.getDate() + 1)) {
    if (matchesCronCriteria(currentDate)) {
      count++;
    }
  }

  return count;
}