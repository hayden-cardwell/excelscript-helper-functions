/** Find a column or set of columns based on a specified column header string or array of header strings and return their indicies. */
function findColumnIndices(
  sheet: ExcelScript.Worksheet,
  colLabels: string | string[]
): number | number[] {
  const findIndex = (label: string) =>
    sheet.getUsedRange().getRow(0).find(label, {}).getColumnIndex();
  if (Array.isArray(colLabels)) {
    return colLabels.map(findIndex);
  } else {
    return findIndex(colLabels);
  }
}

/** Find a column or set of columns based on a specified column header string or array of header strings and return the whole column.
 * Requires findColumnIndicies(). */
function findAndReturnColumns(
  sheet: ExcelScript.Worksheet,
  colLabels: string | string[]
): ExcelScript.Range[] {
  const indicies = findColumnIndices(sheet, colLabels);
  if (Array.isArray(indicies)) {
    return indicies.map((index) => sheet.getRange().getColumn(index));
  } else {
    return sheet.getRange().getColumn(indicies);
  }
}

/** Filter a column based on a supplied array of strings, get the count of filtered items,
 * and optionally clear the filter. Requires findColumnIndicies().*/
function filterAndGetCount(
  sheet: ExcelScript.Worksheet,
  colLabel: string,
  filterValues: string[],
  clearColFilter: boolean = false
): number {
  sheet
    .getAutoFilter()
    .apply(
      sheet.getAutoFilter().getRange(),
      findColumnIndices(sheet, colLabel),
      {
        filterOn: ExcelScript.FilterOn.values,
        values: filterValues,
      }
    );
  let metric = sheet.getUsedRange().getVisibleView().getRowCount() - 1;
  if (clearColFilter) {
    sheet
      .getAutoFilter()
      .clearColumnCriteria(findColumnIndices(sheet, colLabel));
  }
  return metric;
}

/** Uses the excel autofilter based on a column number and an array of strings to filter on */
function applyCustomFilter(
  sheet: ExcelScript.Worksheet,
  colIndex: number,
  values: string[]
) {
  sheet.getAutoFilter().apply(sheet.getAutoFilter().getRange(), colIndex, {
    filterOn: ExcelScript.FilterOn.values,
    values: values,
  });
}

/** After filtering, set visible cells within a given column to a provided value. */
function setVisibleFilteredCells(
  sheet: ExcelScript.Worksheet,
  colIndex: number,
  value: string
) {
  const range = sheet
    .getUsedRange()
    .getVisibleView()
    .getRange()
    .getColumn(colIndex)
    .getOffsetRange(1, 0)
    .getResizedRange(-1, 0);
  range.setValue(value);
}

/** Copy visible rows from the current sheet to the target sheet. */
function copyVisibleRows(
  sourceSheet: ExcelScript.Worksheet,
  targetSheet: ExcelScript.Worksheet
) {
  let visibleFilteredCells = sourceSheet.getRange().getVisibleView();
  targetSheet
    .getRange("A1")
    .copyFrom(
      visibleFilteredCells.getRange(),
      ExcelScript.RangeCopyType.values
    );
}

/** Delete multiple rows based on an array of column labels.
 * Requires findColumnIndices().*/
function deleteColumns(sheet: ExcelScript.Worksheet, colLabels: string[]) {
  const indicies = findColumnIndices(sheet, colLabels) as number[];
  for (const index of indicies) {
    sheet
      .getRange()
      .getColumn(index)
      .delete(ExcelScript.DeleteShiftDirection.left);
  }
}

/** Get the number of the last used row of the sheet. */
function getLastRowValue(sheet: ExcelScript.Worksheet) {
  const addressMatch = sheet
    .getUsedRange()
    .getLastRow()
    .getAddress()
    .match("![A-Z](.*):");
  return addressMatch ? addressMatch[1] : null;
}

/** Fill a column based on what's in the 2nd row of that column. */
function autoFillColumn(
  sheet: ExcelScript.Worksheet,
  col: ExcelScript.Range,
  lastRow: string
) {
  const startCell = col.getRow(1).getAddress().match("!(.*)")[1];
  const endCell = `!${col.getColumnLetter()}${lastRow}`;
  sheet
    .getRange(startCell)
    .autoFill(endCell, ExcelScript.AutoFillType.fillDefault);
}

/** Get all the column labels as an array. */
function getColLabels(sheet: ExcelScript.Worksheet) {
  return sheet.getUsedRange().getRow(0).getValues()[0];
}
