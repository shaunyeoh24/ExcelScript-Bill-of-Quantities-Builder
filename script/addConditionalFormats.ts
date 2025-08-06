class ConditionalFormattingTheme {
// This class encapsulates conditional formatting properties and allows them to be applied to a specified range.

  constructor(
    public readonly name: string,
    private readonly fillColor: string,
    private readonly fontColor: string,
    private readonly setBold: boolean,
    private readonly formulaGenerator: (row: number) => string
  ) {}

  public applyFormatting(workbook: ExcelScript.Workbook, rangeText: string, dataTopRow: number) {
    
    const sheet = workbook.getActiveWorksheet();
    const range = sheet.getRange(rangeText);

    // 01 - Initialize conditional format
    const cfProperties = range.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();

    // 02 - Apply Rule
    cfProperties.getRule().setFormula(this.formulaGenerator(dataTopRow));

    // 03 - Apply Format
    cfProperties.getFormat().getFill().setColor(this.fillColor);
    cfProperties.getFormat().getFont().setColor(this.fontColor);
    cfProperties.getFormat().getFont().setBold(this.setBold);
  }
}

function main(workbook: ExcelScript.Workbook) {

  const sheet = workbook.getActiveWorksheet();

  // 01 - Get data range
  let {dataTopRow, dataBottomRow} = transformTableToActivityObjects(sheet, 9);

  const rangeAddress = `B${dataTopRow}:J${dataBottomRow}`

  const range = sheet.getRange(rangeAddress);

  range.clearAllConditionalFormats();

  // 02 - Define theme color conditional formatting (default blue)
  const themePalette: ConditionalFormattingTheme[] = [
    
    new ConditionalFormattingTheme(
      "Level 0",
      "acb9ca",
      "000000",
      true,
      (row: number) => `=LEN($B${row})-LEN(SUBSTITUTE($B${row},"-","")) = 0`
    ),

    new ConditionalFormattingTheme(
      "Level 1",
      "d6dce4",
      "000000",
      true,
      (dataTopRow: number) => `=LEN($B${dataTopRow})-LEN(SUBSTITUTE($B${dataTopRow},"-","")) = 1`
    ),

    new ConditionalFormattingTheme(
      "Level 2",
      "f0f2f5",
      "000000",
      false,
      (dataTopRow: number) => `=LEN($B${dataTopRow})-LEN(SUBSTITUTE($B${dataTopRow},"-","")) = 2`
    ),

    new ConditionalFormattingTheme(
      "Level 3",
      "ffffff",
      "000000",
      false,
      (dataTopRow: number) => `=LEN($B${dataTopRow})-LEN(SUBSTITUTE($B${dataTopRow},"-","")) = 3`
    ),

    new ConditionalFormattingTheme(
      "Level 4",
      "ffffff",
      "44546a",
      false,
      (dataTopRow: number) => `=LEN($B${dataTopRow})-LEN(SUBSTITUTE($B${dataTopRow},"-","")) > 3`
    ),

  ];

  // 03 - Apply theme color conditional formattings to range
  for (const theme of themePalette) {
    theme.applyFormatting(workbook, rangeAddress, dataTopRow);
  }

  // 04 - Apply lighten text if empty conditional format
  lightenTextConditionalFormatting(workbook, dataTopRow, dataBottomRow)
}


/**
 * Applies conditional formatting to dim rows with no quantity data.
 * 
 * Affects rows where column G is 0, empty (""), or a dash ("-").
 * Highlights text using the specified color across columns B to J.
 *
 * @param {ExcelScript.Workbook} workbook - The Excel workbook context.
 * @param {number} dataTopRow - First row of the data range (after header).
 * @param {number} dataBottomRow - Last row of the data range.
 * @param {string} [textColor="d0cece"] - Font color to apply if condition is met.
 */
function lightenTextConditionalFormatting(
  workbook: ExcelScript.Workbook,
  dataTopRow: number,
  dataBottomRow: number,
  textColor: string = "d0cece",
): void {
  const sheet = workbook.getActiveWorksheet();

  // 01 - Define the target range: columns B to J for all relevant rows
  const range = sheet.getRange(`B${dataTopRow}:J${dataBottomRow}`);

  // 02 - Remove existing conditional formats to avoid stacking
  range.clearAllConditionalFormats();

  // 03 - Create new custom conditional format object
  const lightenTextConditionalFormat = range
    .addConditionalFormat(ExcelScript.ConditionalFormatType.custom)
    .getCustom();

  // 04 - Set formula: triggers if quantity in column G is 0, blank, or "-"
  lightenTextConditionalFormat.getRule().setFormula(
    `=OR($G${dataTopRow} = 0, $G${dataTopRow} = "", $G${dataTopRow} = "-")`
  );

  // 05 - Set font color if condition is true
  lightenTextConditionalFormat.getFormat().getFont().setColor(textColor);
}


/**
 * Transforms a structured Excel activity table into an array of hierarchical activity objects.
 *
 * Scans a worksheet starting from a given header row, determines the data range dynamically,
 * and builds a structured array of objects containing item codes, quantities, rates, and calculated hierarchy levels.
 *
 * Automatically identifies parent-child relationships based on itemCode format (e.g., "1", "1-1", "1-1-1")
 * and strips rate values from parent-level entries.
 *
 * @param {ExcelScript.Worksheet} sheet - The worksheet containing the activity table.
 * @param {number} tableHeaderRow - The row number where the table headers are located (1-based index).
 * @returns {{
 *   activityObjectsArray: {
 *     rowNumber: number,
 *     itemCode: string,
 *     activityName: string,
 *     quantity: (number | string),
 *     unit: string,
 *     rate: (number | string),
 *     cost: string,
 *     hierarchyLevel: number,
 *     hasChild: boolean,
 *   }[],
 *   dataTopRow: number,
 *   dataBottomRow: number,
 * }} Object containing the parsed activity data, top and bottom data row indices.
 */

function transformTableToActivityObjects(
  sheet: ExcelScript.Worksheet,
  tableHeaderRow: number
): {
  activityObjectsArray: {
    rowNumber: number,
    itemCode: string,
    activityName: string,
    quantity: (number | string),
    unit: string,
    rate: (number | string),
    cost: string,
    hierarchyLevel: number,
    hasChild: boolean,
  }[],
  dataTopRow: number,
  dataBottomRow: number,
} {

  // 01 - Determine the vertical bounds of the data table based on column B
  const bottomTotalsRow = sheet.getRange(`B${tableHeaderRow}`)
    .getRangeEdge(ExcelScript.KeyboardDirection.down)
    .getRowIndex() + 1;

  const dataTopRow = tableHeaderRow + 1;
  const dataBottomRow = bottomTotalsRow - 1;

  // 02 - Exit early if no data rows are found below the header
  if (dataBottomRow < dataTopRow) {
    console.log("No activity rows found.");
    return {
      activityObjectsArray: [],
      dataTopRow,
      dataBottomRow
    };
  }

  // 03 - Extract cell formulas from columns B to J for the detected data range
  const tableData2D = sheet.getRange(`B${dataTopRow}:J${dataBottomRow}`).getFormulas();

  // 04 - Convert each row of formulas into an activity object with metadata
  const activityObjectsArray = tableData2D.map((activity, i) => {
    const itemCode = activity[0] as string;
    const activityName = activity[1] as string;
    const quantity = activity[5] as (number | string);
    const unit = activity[6] as string;
    const rate = activity[7] as (number | string);
    const cost = activity[8] as string;

    const rowNumber = dataTopRow + i;
    const hierarchyLevel = (itemCode.match(/-/g) || []).length;

    return {
      rowNumber,
      itemCode,
      activityName,
      quantity,
      unit,
      rate,
      cost,
      hierarchyLevel,
      hasChild: false
    };
  });

  // 05 - Identify parent-child relationships based on itemCode structure
  for (let i = 0; i < activityObjectsArray.length - 1; i++) {
    const current = activityObjectsArray[i];
    const next = activityObjectsArray[i + 1];

    if (
      next.itemCode.startsWith(current.itemCode + "-") &&
      next.hierarchyLevel === current.hierarchyLevel + 1
    ) {
      current.hasChild = true;
    }
  }

  // 06 - Clear rate values for items that act as parents in the hierarchy

  activityObjectsArray.forEach((activity) => {
    if (activity.hasChild) {
      activity.rate = "";
    }
  });

  return { activityObjectsArray, dataTopRow, dataBottomRow };
}







