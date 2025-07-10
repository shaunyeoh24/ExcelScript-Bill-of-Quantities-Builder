function main(workbook: ExcelScript.Workbook) {

    // 01 - Get the currently active worksheet
    let sheet = workbook.getActiveWorksheet();

    // 02 - Extract structured activity data and table bounds from worksheet
    const { activityObjectsArray, dataTopRow, dataBottomRow } = transformTableToActivityObjects(sheet, 9);

    // 03 - Update quantity field with "[qty]" placeholder for all non-parent activities
    activityObjectsArray.forEach((activity) => {
        if (!activity.hasChild) {
            activity.quantity = "[qty]";
        }
    });

    // 04 - Build 2D array of updated quantity values for column G
    const activityQuantities = activityObjectsArray.map(obj => [obj.quantity]);

    // 05 - Apply quantity updates to the worksheet (column G only)
    const tableDataRange = sheet.getRange(`G${dataTopRow}:G${dataBottomRow}`);
    tableDataRange.setFormulas(activityQuantities as string[][]);

    // 06 - Log completion status
    console.log("Reset Quantities: Completed successfully.");
}


// ===== HELPER 01: TRANSFORM TABLE TO ACTIVITY OBJECT ===== //

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
