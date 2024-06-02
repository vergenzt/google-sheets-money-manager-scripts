function test() {
  console.log(RRULE_TOSTRING("every 3 months starting 2024-02-27"));
}

function new_rrule(ruletext) {
  if (!ruletext) throw new Error("No ruletext provided");
  let [base, dtstart_text] = ruletext.split(" from ");
  let dtstart = dtstart_text ? new Date(dtstart_text) : undefined;
  return new rrule.RRule({ dtstart, ...rrule.RRule.parseText(base) });
}

/**
 * @customfunction
 */
function RRULE_TOSTRING(ruletext) {
  return new_rrule(ruletext).toString();
}

/**
 * @customfunction
 */
function RRULE_NEXT(ruletext, after=new Date()) {
  return new_rrule(ruletext).after(after);
}

/**
 * Yields an array of upcoming dates for the given rules array, sorted by date.
 * @customfunction
 */
function RRULES_UPCOMING(ruletexts, date_from, date_to, extra) {
  return ruletexts
    .flatMap(([ruletext], i) => {
      try {
        let rule = new_rrule(ruletext);
        let upcoming_dates = rule.between(date_from, date_to);
        return upcoming_dates.map(date => [date, ...extra[i]]);
      } catch {
        return [];
      }
    })
    .sort((a, b) => a[0] - b[0]);
}

/**
 * Appends rows to "Upcoming" sheet based on the "Series" sheet.
 */
function GenerateUpcoming() {
  /* pseudocode:
  for each row in Series:
    generate upcoming dates
    for each upcoming date according to the Series' recurrence rule:
      append 
  */
}