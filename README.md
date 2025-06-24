# excelscript-bill-of-quantities-builder
An ExcelScript-powered tool for dynamically building, organizing, and automating hierarchical Bills of Quantities (BQ) in Excel — featuring parent-child structuring, formula auto-generation, and customizable item insertion logic.

---

This ExcelScript automates the management of hierarchical activity tables in Excel—specifically for civil, construction, or project-based scopes that involve structured item codes (e.g., A-1-2). It supports automated child row insertion, styling, and dynamic formula updates for quantity, unit, rate, and cost fields.

✨ Features
Insert Child Item Row

Dynamically generates and inserts a new child row under the selected item.

Maintains item code hierarchy (e.g., A-1 → A-1-1, A-1-2, etc.).

Applies structured formatting and merged cells per level (0–3).

Auto-indents and color-codes by depth.

Auto-Update Activity Formulas

Recalculates quantities, units, rates, and cost formulas for the entire activity table.

Detects parent-child relationships to roll up costs using SUM() formulas.

Applies intelligent fallbacks and placeholders for manual entries.

📁 File Structure (Key Functions)
Function	Purpose
main()	Entry point. Inserts row and triggers formula update.
insertChildItemRow()	Inserts a child row based on active cell and hierarchy rules.
updateActivityRowFormulas()	Parses activity rows and updates rate and cost formulas accordingly.
computeNewChildItemCodeAndInsertionRow()	Calculates next child code and insertion point.
formatRowByHierarchyLevel()	Applies visual formatting and cell merging based on depth.

🔢 Hierarchy Levels
Level	Example Code	Formatting Description
0	A	Section Header (Bold, Dark Fill)
1	A-1	Parent Item (Gray Fill)
2	A-1-1	Activity (White Fill)
3	A-1-1-1	Sub-Activity (Blue Font)

📌 Usage
Trigger the Script
Select a cell in Column B containing a valid item code and run the script.
It inserts a new child item directly beneath it.

Auto-Recalculate Formulas
After any structural change, formulas for quantity, unit, rate, and cost are recalculated.

Edit Safely
Only modify columns B, C–F, G–I manually. Column J (Cost) is auto-derived and should not be edited directly.

⚠️ Notes & Assumptions
Assumes item codes in Column B follow a hyphenated format (e.g., A, A-1, A-1-2).

Table starts at a configurable tableHeaderRow (default: 9).

The bottom of the table is defined by a grand total row (non-empty in Column B).
