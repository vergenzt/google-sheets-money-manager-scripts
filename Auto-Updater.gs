
function autoUpdateTransactions() {
  let transactions = TransactionsSheet();
  let txHead = transactions.getRange('1:1') //.getValues()[0];
}

function installAutoUpdate() {
  ScriptApp
    .newTrigger('autoUpdateTransactions')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}
