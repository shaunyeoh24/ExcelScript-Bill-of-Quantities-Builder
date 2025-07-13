function main(workbook: ExcelScript.Workbook) {

  const sheet = workbook.getActiveWorksheet();

  insertChildItemRow(workbook);
  updateActivityRowFormulas(sheet, 9);

}



// ====== PROGRAM FUNCTIONS ===== //

/**
 * Inserts a sibling item row below the currently active cell in the worksheet.
 * The inserted row is formatted based on the item's hierarchy level determined from the item code in column B.
 * Supports 0 (parent) to 3 (child) levels.
 * 
 * @param workbook The ExcelScript workbook object.
 */
function insertChildItemRow(workbook: ExcelScript.Workbook): void {
  const sheet = workbook.getActiveWorksheet();
  const activeCell = workbook.getActiveCell();
  const activeCellRowNumber = activeCell.getRowIndex() + 1;

  console.log(`Active Cell Row: ${activeCellRowNumber}`)

  // Get the parent item code properties and compute child item insertion properties
  const { newChildInsertionRow, newChildItemCode, newChildHierarchyLevel } = computeNewChildItemCodeAndInsertionRow(sheet, activeCellRowNumber);

  console.log({ newChildInsertionRow, newChildItemCode, newChildHierarchyLevel });

  // Define headers for levels 0–3 (columns B–J, with C–F being merged)
  const headersArray: string[][][] = [
    [[newChildItemCode, "[SECTION X - PARENT ITEM TEMPLATE]", "", "", "", "[qty]", "[unit]", "[rate]", "[insert formula]"]],
    [[newChildItemCode, "[Insert Activity/Item]", "", "", "", "[qty]", "[unit]", "[rate]", "[insert formula]"]],
    [[newChildItemCode, "[Insert Activity/Item]", "", "", "", "[qty]", "[unit]", "[rate]", "[insert formula]"]],
    [[newChildItemCode, "[Insert Activity/Item]", "", "", "", "[qty]", "[unit]", "[rate]", "[insert formula]"]]
  ];

  // Insert blank row below the active cell, shifting existing cells down
  sheet.getRange(`${newChildInsertionRow}:${newChildInsertionRow}`)
    .insert(ExcelScript.InsertShiftDirection.down);

  console.log(headersArray[newChildHierarchyLevel]);

  // Insert header values into new row (starting at column B, spanning to J)
  const targetRange = sheet.getRange(`B${newChildInsertionRow}`).getResizedRange(0, 8);
  targetRange.setValues(headersArray[Math.min(newChildHierarchyLevel, 3)]);

  // Apply formatting based on hierarchy level
  formatRowByHierarchyLevel(workbook, newChildHierarchyLevel, newChildInsertionRow);

  // Reset active cell to the Column B of the new row
  // sheet.getRange(`B${newChildInsertionRow}`).select();

  // Log function completion
  console.log(`✅ Inserted new child item at row ${newChildInsertionRow} with hierarchy level ${newChildHierarchyLevel}`);
}




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

        } else {

            // Assign placeholders or cost formula for leaf items
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




// ====== HELPER FUNCTIONS ===== //

/**
 * Applies row formatting (fill, font, alignment, indent, merge) based on the item's hierarchy level.
 *
 * @param workbook The ExcelScript workbook object.
 * @param hierarchyLevel The hierarchy level (0–3) used to determine formatting.
 * @param rowNumberToFormat The row number (1-based) to apply the formatting to.
 */
function formatRowByHierarchyLevel(
  workbook: ExcelScript.Workbook,
  hierarchyLevel: number,
  rowNumberToFormat: number
): void {
  const sheet = workbook.getActiveWorksheet();

  const indentLevel = hierarchyLevel + 1;
  const rowRange = sheet.getRange(`B${rowNumberToFormat}:J${rowNumberToFormat}`);
  const itemNameRange = sheet.getRange(`C${rowNumberToFormat}:F${rowNumberToFormat}`);

  // Style settings by hierarchy level (0 = parent, 3 = deepest child)
  const rowFontAndFillStyles = [
    { fill: "acb9ca", fontSize: 11, color: "000000", bold: true, },
    { fill: "d6dce4", fontSize: 9, color: "000000", bold: true },
    { fill: "ffffff", fontSize: 8, color: "000000", bold: false },
    { fill: "ffffff", fontSize: 8, color: "44546a", bold: false }
  ];

  // Cap the style to a maximum hierarchy level of 3
  const style = rowFontAndFillStyles[Math.min(hierarchyLevel, 3)];

  // Apply row-wide formatting
  const format = rowRange.getFormat();
  format.getFill().setColor(style.fill);
  format.getFont().setSize(style.fontSize);
  format.getFont().setColor(style.color);
  format.getFont().setBold(style.bold);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  // Column-specific horizontal alignment and indenting
  const alignSpecs = [
    { column: "B", hAlign: ExcelScript.HorizontalAlignment.center, indent: 0 },
    { column: "C", hAlign: ExcelScript.HorizontalAlignment.left, indent: indentLevel },
    { column: "G", hAlign: ExcelScript.HorizontalAlignment.center, indent: 0 },
    { column: "H", hAlign: ExcelScript.HorizontalAlignment.center, indent: 0 },
    { column: "I", hAlign: ExcelScript.HorizontalAlignment.right, indent: 1 },
    { column: "J", hAlign: ExcelScript.HorizontalAlignment.right, indent: 1 }
  ];

  alignSpecs.forEach(spec => {
    const cell = sheet.getRange(`${spec.column}${rowNumberToFormat}`);
    const cellFormat = cell.getFormat();
    cellFormat.setHorizontalAlignment(spec.hAlign as ExcelScript.HorizontalAlignment);
    cellFormat.setIndentLevel(spec.indent);
  });

  // Merge description columns (C–F)
  itemNameRange.merge();

  // Number format for price-related columns
  const pricesRange = sheet.getRange(`I${rowNumberToFormat}:J${rowNumberToFormat}`)
  pricesRange.setNumberFormatLocal("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
}


/**
 * Retrieves all item codes from a table starting at a specific row in column B.
 * Delegates to a general-purpose helper that scans downward from the given row.
 *
 * @param sheet The ExcelScript worksheet object.
 * @param itemCodeStartRow The row number (1-based) to start scanning from.
 * @returns An array of objects, each containing the row number and corresponding item code string.
 */
function getAllItemCodesInTable(
  sheet: ExcelScript.Worksheet,
  itemCodeStartRow: number
): { rowNumber: number, itemCode: string }[] {
  return getAllItemCodesFromGivenRow(sheet, "B", itemCodeStartRow);
}


/**
 * Retrieves all non-empty item codes starting from a given row in a specified column,
 * scanning downward until an empty cell is encountered.
 *
 * @param sheet The ExcelScript worksheet object.
 * @param columnLetter The column letter (e.g. "B") that contains item codes.
 * @param startRow The row number (1-based) to start scanning from.
 * @returns An array of objects containing row numbers and corresponding item codes.
 */
function getAllItemCodesFromGivenRow(
  sheet: ExcelScript.Worksheet,
  columnLetter: string,
  startRow: number
): { rowNumber: number, itemCode: string }[] {

  // Get all cell values from the start row downward in the target column
  const values2D = sheet
    .getRange(`${columnLetter}${startRow}`)
    .getExtendedRange(ExcelScript.KeyboardDirection.down)
    .getValues();

  const itemCodes: { rowNumber: number, itemCode: string }[] = [];

  for (let i = 0; i < values2D.length; i++) {
    const rowNumber = startRow + i;
    const itemCode = values2D[i][0] as string;
    itemCodes.push({ rowNumber, itemCode });
  }

  return itemCodes;
}

/**
 * Returns all child activity items that belong to the given parent item row,
 * based on the item code hierarchy in column B.
 * 
 * Child items are determined by:
 * - Having item codes that start with the parent code followed by a hyphen.
 * - Having a deeper hierarchy level (more hyphens) than the parent.
 *
 * @param sheet - The worksheet object.
 * @param rowNumber - The row number (1-based) of the parent item.
 * @returns An array of objects, each containing the row number and item code of a child activity.
 */
function getAllChildActivityItemCodes(
  sheet: ExcelScript.Worksheet,
  rowNumber: number
): { rowNumber: number, itemCode: string }[] {

  // Get parent item code and compute parent level from column B
  const parentCode = sheet.getRange(`B${rowNumber}`).getValue() as string;
  const parentLevel: number = (parentCode.match(/-/g) || []).length;

  // Get all item codes below the parent row
  const itemCodesBelow = getAllItemCodesFromGivenRow(sheet, "B", rowNumber + 1);

  // Accumulator for child items
  const childrenActivities: { rowNumber: number, itemCode: string }[] = [];

  // Loop through each item code below
  itemCodesBelow.forEach(item => {
    const itemLevel = (item.itemCode.match(/-/g) || []).length;

    // Must start with parentCode + "-" and be at a deeper level
    if (
      item.itemCode.startsWith(`${parentCode}-`) &&
      itemLevel > parentLevel
    ) {
      childrenActivities.push(item);
    }
  });

  return childrenActivities;
}




/**
 * Computes the new child item code and the row where it should be inserted,
 * based on a selected parent row in the Excel worksheet.
 *
 * @param sheet - The ExcelScript worksheet containing the item codes.
 * @param selectedParentRow - The row number (1-based) of the selected parent item.
 * @returns An object containing:
 *  - newChildInsertionRow: the row number to insert the new child,
 *  - newChildItemCode: the generated item code for the new child,
 *  - newChildHierarchyLevel: the depth level in the item hierarchy.
 */
function computeNewChildItemCodeAndInsertionRow(
  sheet: ExcelScript.Worksheet,
  selectedParentRow: number
): { newChildInsertionRow: number, newChildItemCode: string, newChildHierarchyLevel: number } {

  // Get the parent item code from column B of the selected row
  const parentItemCode = sheet.getRange(`B${selectedParentRow}`).getValue() as string;

  // Determine hierarchy depth from number of hyphens (e.g., A-1-2 = 2)
  const parentHierarchyLevel: number = (parentItemCode.match(/-/g) || []).length;

  // Retrieve all child activity items under the selected parent
  const childItems = getAllChildActivityItemCodes(sheet, selectedParentRow);

  // CASE 1: No existing children — create the first child at next row
  if (childItems.length === 0) {
    const newChildInsertionRow: number = selectedParentRow + 1;
    const newChildItemCode: string = parentItemCode + "-1";
    const newChildHierarchyLevel: number = parentHierarchyLevel + 1;

    return { newChildInsertionRow, newChildItemCode, newChildHierarchyLevel };

  } else {
    // CASE 2: Existing children — find the last one with the correct hierarchy level
    const matchingLevelChildren: { rowNumber: number, itemCode: string }[] = [];

    childItems.forEach((item) => {
      const itemHierarchyLevel = item.itemCode.match(/-/g).length;
      if (itemHierarchyLevel === parentHierarchyLevel + 1) {
        matchingLevelChildren.push(item);
      }
    });

    // Take the last matching child and increment its item code
    const lastImmediateChild = matchingLevelChildren[matchingLevelChildren.length - 1];
    const lastAbsoluteChild = childItems[childItems.length - 1];

    const newChildInsertionRow: number = lastAbsoluteChild.rowNumber + 1;
    const newChildItemCode: string = incrementLastItemCodeSegment(lastImmediateChild.itemCode);
    const newChildHierarchyLevel: number = parentHierarchyLevel + 1;

    return { newChildInsertionRow, newChildItemCode, newChildHierarchyLevel };
  }
}



/**
 * Increments the final numeric segment of a hyphen-delimited item code string.
 * Example: 'A-1-3' → 'A-1-4'
 *
 * @param itemCode - The item code string to increment (e.g., "A-1-2").
 * @returns The incremented item code string.
 * @throws If the last segment is not a valid number.
 */
function incrementLastItemCodeSegment(itemCode: string): string {

  const segments = itemCode.split("-");
  const lastSegment = Number(segments[segments.length - 1]);

  if (isNaN(lastSegment)) {
    throw new Error(`Invalid numeric segment in code: ${itemCode}`);
  }

  // Replace last segment with incremented value
  segments[segments.length - 1] = String(lastSegment + 1);

  return segments.join("-");
}




