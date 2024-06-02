function testTxRulesApply() {
  let sheet = SpreadsheetApp.getActive();
  let descs = sheet.getRange("Transactions!T2:T").getValues();
  let rules = sheet.getRangeByName("Rules").getValues();
  let column = "Description";
  TX_RULES_APPLY(descs, rules, column);
}


/**
 * @customfunction
 * @param {[[string]]} fullDescs transaction descriptions column
 * @param {[[string]]} rulesTable the rules sheet
 * @param {string} column the name of the rule result column to select
 */
function TX_RULES_APPLY(fullDescs, rulesTable, column) {
  let rulesHead = rulesTable[0];
  let rules = rulesTable.slice(1).map(rulesRow => 
    Object.fromEntries(rulesRow.map((value, i) => [rulesHead[i], value]))
  );

  rules.forEach(rule => {
    rule.Regex = new RegExp(
      (
        rule.Words
        .replace("([A-Z]{2})([a-z])", "$1 $2")
        .split(" ")
        .map(word => `\\b${word}\\b`)
        .join(".*")
      ),
      "i"
    );
  });

  let wordsOf = fullDesc => (
    fullDesc && fullDesc
    .replace(/([A-Z]{2})([a-z])/, "$1 $2")
    .replace(/([a-z])([A-Z])/, "$1 $2")
    .replace(/[xX]{4,}/, " ")
  );

  return fullDescs.map(([fullDesc]) => {
    let words = wordsOf(fullDesc);
    let rule = rules.find(rule => rule.Regex.test(words));
    return rule ? rule[column] : '';
  });
}