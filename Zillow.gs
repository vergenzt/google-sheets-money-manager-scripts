function GET_HOME_VALUE() {
  let contents = UrlFetchApp.fetch("https://www.zillow.com/homes/4580-Bannons-Walk-Ct_rb/44576637_zpid/").getContentText();
  let match = /\bprice.{1,3}\b(\d{6})\b/.exec(contents);
  console.log(match[1]);
  return match[1];
}