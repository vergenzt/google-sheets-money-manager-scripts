




function expensesCalId() {
  return PropertiesService.getScriptProperties().getProperty("EXPENSES_CALENDAR_ID");
}

function stringProp(string, prop) {
  let re = new RegExp(`^{prop}: (.*)`, 'm');
  let match = re.exec(string);
  return match ? match[1] : '';
}

function GET_SERIES() {
  let events = Calendar.Events.list(expensesCalId());
  let ret = [
    ["Label", "Payee", "Amount", "Category", "Recurrence"],
    ...events.items.map(event => [
      event.summary,
      stringProp(event.description, 'Payee'),
      stringProp(event.description, 'Amount'),
      stringProp(event.description, 'Category'),
      event.recurrence,
    ]),
  ];
  return ret;
}
