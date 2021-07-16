const expectedHoursPerDay = 8;

const mainSheet = SpreadsheetApp.getActive().getSheetByName('Today');
const monthSheet = SpreadsheetApp.getActive().getSheetByName('Month');

const mainSheetSumRow = 1;
const mainSheetSumColumn = 6;
const startLabel = 'Start'

const firstMothValueRow = 2;

const balanceCell = {
  row: 1,
  column: 7
}

function logStart() {
  if (isStarted()) {
    return;
  }
  mainSheet.appendRow([startLabel, new Date(), `=IF(C${mainSheet.getLastRow() + 2}=""; NOW() - B${mainSheet.getLastRow() + 1}; "")`]);
}

function logEnd() {
  if (!isStarted()) {
    return;
  }
  const formula = `=B${mainSheet.getLastRow() + 1}-B${mainSheet.getLastRow()}`
  mainSheet.appendRow(['Stop', new Date(), formula]);
}

function isStarted() {
  if (mainSheet.getLastRow() < 1) {
    return false;
  }
  const lastAction = mainSheet.getRange(mainSheet.getLastRow(), 1).getValue();
  return lastAction == startLabel || lastAction == ''
}

function getFirstDate() {
  const value = mainSheet.getRange(2, 2).getValue();
  return value && new Date(value.toString()).toISOString().split('T')[0]
}

function getMainSum() {
  return mainSheet.getRange(mainSheetSumRow, mainSheetSumColumn).getValue();
}

function submit() {
  const date = getFirstDate();
  if (date === undefined || date === '') {
    return;
  }
  logEnd();
  monthSheet.appendRow([date, getMainSum()]);
  clearToday();
  updateBalance();
}

function clearToday() {
  mainSheet.getRange(2, 1, 1000, 3).clear();
}

function clearMonth() {
  getMonthRange().clear();
  getBalanceCell().clear();
}

function updateBalance() {
  getBalanceCell().setValue(
    secondsToString(getWorkedSeconds() - getExpectedSeconds())
  );
}

function getWorkedSeconds() {
  const lastRow = monthSheet.getLastRow();
  const firstRow = 2;
  const workData = monthSheet.getRange(firstRow, 2, lastRow - firstRow + 1, 1);
  return workData.getValues().flat().map(cellToSeconds).reduce((prev, curr) => prev + curr);
}

function cellToSeconds(cellValue) {
  const date = new Date(cellValue);
  return date.getSeconds() + date.getMinutes() * 60 + date.getHours() * 3600;
}

function getMonthRange() {
  const lastRow = monthSheet.getLastRow();
  const numRows = lastRow - firstMothValueRow + 1;
  if (numRows < 1) {
    return { clear: () => { } };
  }
  return monthSheet.getRange(firstMothValueRow, 1, lastRow - firstMothValueRow + 1, 2);
}

function getBalanceCell() {
  return monthSheet.getRange(
    balanceCell.row, balanceCell.column
  )
}

function getExpectedSeconds() {
  return getMonthRange().getNumRows() * expectedHoursPerDay * 3600;
}

function secondsToString(asSeconds) {
  const sign = asSeconds > 0 ? '' : '-'
  asSeconds = Math.abs(asSeconds);
  const hours = Math.trunc(asSeconds / 3600);
  const minutes = Math.round((asSeconds - hours * 3600) / 60);
  const seconds = Math.round((asSeconds - hours * 3600 - minutes * 60));
  return `${sign}${hours}:${minutes}:${seconds}`;
}
