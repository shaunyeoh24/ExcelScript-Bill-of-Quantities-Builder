/**
 * ExcelScript Module: Delete Selected Activities and Refresh
 *
 * Automates:
 * - Deleting user-selected activity rows
 * - Reindexing item codes hierarchically
 * - Reapplying formulas based on hierarchical structure
 *
 * @entryPoint `main(workbook: ExcelScript.Workbook)`
 *
 * @assumptions:
 * - Table starts at row 9 with item codes in Column B
 * - Hierarchy denoted by hyphenated codes (e.g., A-1, A-1-1)
 *
 * @dependencies:
 * - refreshActivityItemCodes
 * - updateActivityRowFormulas
 * - transformTableToActivityObjects
 * - reindexItemCodes
 */

// ===== MAIN FUNCTION ===== //
function main(workbook: ExcelScript.Workbook) {

  deleteSelectedActivitiesAndRefresh(workbook);

}

// ===== PROGRAM FUNCTION 01 - DELETE SELECTED ACTIVITIES AND REFRESH ===== //
/**
 * Deletes the currently selected activity row(s) from the active worksheet,
 * then refreshes item codes and updates related formulas.
 *
 * @param {ExcelScript.Workbook} workbook - The workbook containing the active worksheet and selected range.
 *
 * @dependencies Dependencies are as follows:
 *   - Requires `refreshActivityItemCodes`
 *   - Requires `updateActivityRowFormulas`
 *   - Relies on item codes in Column B
 *   - Data range must be contiguous
 */
function deleteSelectedActivitiesAndRefresh(workbook: ExcelScript.Workbook) {

  // 01 - Get the active cell and worksheet
  const sheet = workbook.getActiveWorksheet();
  const activeRange = workbook.getSelectedRange();

  const activeRangeTopRow = activeRange.getRowIndex() + 1;
  const activeRangeBottomRow = activeRangeTopRow + activeRange.getRowCount() - 1;

  const {dataTopRow, dataBottomRow} = transformTableToActivityObjects(sheet, 9);

  // 02 - Validate deletion range
  if (activeRangeTopRow < dataTopRow || activeRangeBottomRow > dataBottomRow) throw new Error("Selected row not deleted - row not within data range!")

  // 03 - Delete selected row(s)
  sheet.getRange(`${activeRangeTopRow}:${activeRangeBottomRow}`).delete(ExcelScript.DeleteShiftDirection.up);

  // 04 - Reindex and reset itemCodes
  refreshActivityItemCodes(workbook, 9);

  // 05 - Update formulas for quantity, unit, rate, cost columns
  updateActivityRowFormulas(sheet, 9);

}

// ===== PROGRAM FUNCTION 02 - REFRESH ACTIVITY ITEM CODES ===== //

/**
 * Refreshes the activity item codes in the currently active worksheet of the workbook.
 * 
 * This function:
 * - Extracts activity data from a structured table based on the provided header row.
 * - Recalculates item codes using `reindexItemCodes()`.
 * - Updates column B with the refreshed item codes for all detected activity rows.
 * 
 * @param {ExcelScript.Workbook} workbook - The Excel workbook containing the active worksheet.
 * @param {number} tableHeaderRow - The row number where the table headers begin (used to identify the table range).
 *
 * @remarks
 * Relies on the helper functions `transformTableToActivityObjects` and `reindexItemCodes`
 * being available in the execution context.
 *
 * @example
 * Refresh item codes for a table starting at header row 9
 * refreshActivityItemCodes(workbook, 9);
 */

function refreshActivityItemCodes(workbook: ExcelScript.Workbook, tableHeaderRow: number): void {

  // 01 - Get the currently active worksheet
  let sheet = workbook.getActiveWorksheet();

  // 02 - Extract structured activity data and table bounds from worksheet
  const { activityObjectsArray, dataTopRow, dataBottomRow } = transformTableToActivityObjects(sheet, tableHeaderRow);


  // 03 - Prepare array of itemCode and hierarchyLevel objects
  const existingItemCodes = activityObjectsArray.map((activity) => {
    return activity.itemCode;
  });

  // 04 - Generate refreshed activity itemCode array 
  const newItemCodes = reindexItemCodes(existingItemCodes);

  // 05 - Generate 2D array and reapply itemCodes onto sheet
  const newItemCodes2D = newItemCodes.map(itemCode => [itemCode]);
  sheet.getRange(`B${dataTopRow}:B${dataBottomRow}`).setFormulas(newItemCodes2D);

}

// ===== PROGRAM FUNCTION 03 - UPDATE ACTIVITY ROW FORMULAS ===== //

/**
 * Updates the activity table in the worksheet by processing hierarchical item codes
 * and applying quantity, unit, rate, and cost formulas accordingly.
 *
 * @param sheet - The Excel worksheet containing the activity table.
 * @param tableHeaderRow - The row number where the table headers begin (e.g. 9).
 * 
 * This function:
 * - Identifies the range of rows containing data under the header.
 * - Parses each row into an object with relevant properties.
 * - Determines hierarchy level and whether an item has children.
 * - Applies formulas or placeholders based on hierarchy structure.
 * - Writes updated values/formulas to columns G to J.
 */
function updateActivityRowFormulas(
  sheet: ExcelScript.Worksheet,
  tableHeaderRow: number,
) {

  // Get bottom grand totals row index
  const bottomTotalsRow = sheet.getRange(`B${tableHeaderRow}`).getRangeEdge(ExcelScript.KeyboardDirection.down).getRowIndex() + 1;

  // Define top and bottom data row bounds
  const dataTopRow = tableHeaderRow + 1;
  const dataBottomRow = bottomTotalsRow - 1;

  // Exit early if no activity rows exist
  if (dataBottomRow === tableHeaderRow) {
    console.log("The table is empty!")
    return [];
  }

  // Extract data range from columns B to J
  const tableData2D = sheet.getRange(`B${dataTopRow}:J${dataBottomRow}`).getFormulas();

  // Map each row into an activity object
  const activityObjectsArray: {
    rowNumber: number,
    itemCode: string,
    quantity: (number | string),
    unit: string,
    rate: (number | string),
    cost: string,
    hierarchyLevel: number,
    hasChild: boolean,
  }[] = tableData2D.map((activity, i) => {

    // Formulas from existing table
    const itemCode = activity[0] as string;
    const quantity = activity[5] as (number | string);
    const unit = activity[6] as string;
    const rate = activity[7] as (number | string);
    const cost = activity[8] as string;

    // Formulas not from table
    const rowNumber = dataTopRow + i as number;
    const hierarchyLevel = (itemCode.match(/-/g) || []).length as number;

    return {
      rowNumber,
      itemCode,
      quantity,
      unit,
      rate,
      cost,
      hierarchyLevel,
      hasChild: false
    };
  });

  // Check and mark each activity that has at least one child
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

  // Update each activity’s fields based on whether it has children
  activityObjectsArray.forEach((activity) => {

    if (activity.hasChild) {

      const preserveQtyAndUnit: boolean =
        activity.unit !== "LS" &&
        activity.quantity !== 1 &&
        activity.quantity !== "[qty]";

      if (preserveQtyAndUnit) {

        // Keep original quantity and unit
        let rateFormula = "=SUM(";

        // Iterate of child activities
        activityObjectsArray.forEach(item => {
          if (
            item.itemCode.startsWith(activity.itemCode + "-") &&
            item.hierarchyLevel === activity.hierarchyLevel + 1
          ) {
            rateFormula += `J${item.rowNumber},`;
          }
        });

        rateFormula = rateFormula.slice(0, -1) + ")";

        activity.rate = rateFormula;
        activity.cost = `=IFERROR(IF(G${activity.rowNumber}="-", "-", G${activity.rowNumber}*I${activity.rowNumber}), "[pending values]")`;

      } else {

        // Aggregate formula logic for parent items

        activity.quantity = 1;
        activity.unit = "LS";

        let costFormula = "=SUM(";

        activityObjectsArray.forEach(item => {
          if (
            item.itemCode.startsWith(activity.itemCode + "-") &&
            item.hierarchyLevel === activity.hierarchyLevel + 1
          ) {
            costFormula += `J${item.rowNumber},`;
          }
        });

        costFormula = costFormula.slice(0, -1) + ")";

        activity.rate = "";
        activity.cost = costFormula;

      }

    } else {

      // Assign placeholders or cost formula for all leaf items
      activity.quantity = activity.quantity === "" ? "[qty]" : activity.quantity;
      activity.unit = activity.unit === "" ? "[unit]" : activity.unit;
      activity.rate = (activity.rate === "" || activity.rate === "#REF!") ? "[rate]" : activity.rate;

      activity.cost = `=IFERROR(IF(G${activity.rowNumber}="-", "-", G${activity.rowNumber}*I${activity.rowNumber}), "[pending values]")`;
    }
  });

  // Prepare data arrays for setting back to sheet
  const formulasArray2D = activityObjectsArray.map(activity => [

    activity.quantity,
    activity.unit,
    activity.rate,
    activity.cost

  ]);

  // Write values and formulas back to worksheet
  sheet.getRange(`G${dataTopRow}:J${dataBottomRow}`).setFormulas(formulasArray2D as string[][]);
}




// ===== HELPER FUNCTION 01 - REINDEX ITEM CODES ===== //

/**
 * Reindexes a list of hierarchical item codes to maintain contiguous numbering for each level.
 * 
 * Example:
 * Input:  ["A", "A-1", "A-1-1", "A-1-3", "A-2", "A-2-1", "A-3", "A-3-2"]
 * Output: ["A", "A-1", "A-1-1", "A-1-2", "A-2", "A-2-1", "A-3", "A-3-1"]
 * 
 * Assumes:
 * - Input is in depth-first traversal order
 * - Hierarchy is represented using hyphen-separated strings (e.g., "A-1-2")
 *
 * @param {string[]} itemCodes - Array of item codes representing a hierarchical structure
 * @returns {string[]} - A new array with reindexed item codes preserving hierarchy
 */

function reindexItemCodes(itemCodes: string[]): string[] {
  // 01 - Initialize the array to hold reindexed item codes
  const reindexed: string[] = [];

  // 02 - Initialize a path stack to build new item codes by hierarchy level
  const pathStack: string[] = [];

  // 03 - Initialize a counter array to track sibling indexes at each level
  const levelCounters: number[] = [];

  // 04 - Iterate through each itemCode in the provided array
  for (let i = 0; i < itemCodes.length; i++) {
    const code = itemCodes[i];

    // 04a - Split the itemCode by '-' to determine its hierarchical level
    const parts = code.split("-");

    // 04b - Determine the current depth level (0-based)
    const level = parts.length - 1;

    // 04c - Truncate the pathStack and levelCounters to match the current depth
    pathStack.length = level;
    levelCounters.length = level + 1;

    // 04d - Handle root-level items (e.g., "A")
    if (level === 0) {
      // 04d (i) - Directly assign the root name to pathStack[0]
      pathStack[0] = parts[0];

      // 04d (ii) - Set root-level counter to 0 (not actively used)
      levelCounters[0] = 0;
    } else {
      // 04e (i) - Initialize or increment the sibling counter at current level
      levelCounters[level] = (levelCounters[level] || 0) + 1;

      // 04e (ii) - Store current index in the pathStack for this level
      pathStack[level] = levelCounters[level].toString();
    }

    // 04f - Rebuild the itemCode string based on the updated pathStack
    const newCode = (level === 0)
      ? pathStack[0]  // 04f (i) - Root level: just use the root name
      : `${pathStack[0]}-${pathStack.slice(1, level + 1).join("-")}`; // 04f (ii) - Join updated hierarchy

    // 04g - Push the newCode into the result array
    reindexed.push(newCode);
  }

  // 05 - Return the reindexed list of item codes
  return reindexed;
}


// ===== HELPER FUNCTION 02 - TRANSFORM TABLE TO ACTIVITY OBJECTS ===== //

/**
 * Transforms a structured table in an Excel worksheet into an array of activity objects,
 * capturing hierarchy and parent-child relationships based on item codes.
 *
 * @param {ExcelScript.Worksheet} sheet - The Excel worksheet containing the table.
 * @param {number} tableHeaderRow - The row number where the table header begins.
 * 
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
 * }} An object containing:
 * - `activityObjectsArray`: Array of activity rows with metadata and hierarchy information.
 * - `dataTopRow`: The first data row index (after header).
 * - `dataBottomRow`: The last data row index (before totals).
 *
 * @example
 * const result = transformTableToActivityObjects(sheet, 4);
 * console.log(result.activityObjectsArray);
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
