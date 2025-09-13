
# NWE Shipping Refund Calculator - Developer Guide

This guide provides a comprehensive technical overview of the PowerShell GUI application, including code structure, logic, and maintenance tips. It is written for junior developers and explains every major part of the code and why it exists.

---

## Application Overview

- Loads order data from CSV
- Calculates shipping totals and refunds per recipient
- Highlights missing/invalid data
- Displays summary statistics and customer stats
- Allows editing and saving results
- Exports purchases and stats to CSV
- Built with Windows Forms for a modern GUI
- Robust error handling and maintainable code structure

---

## CSV Column Mapping

The script uses variables to reference CSV column names. Update these variables if your CSV headers change. **They must match your CSV file's column names exactly.**

| Script Variable         | Expected Column Name in CSV      | Description                                      |
|------------------------|----------------------------------|--------------------------------------------------|
| $colOrder              | Order #                          | The order number for each purchase               |
| $colItemName           | Item Name                        | The name of the item ordered                     |
| $colRecipient          | Recipient                        | The customer or recipient of the order           |
| $colQuantity           | Quantity                         | Number of items in the order                     |
| $colOrderTotal         | Order Total                      | Total cost of the order (excluding shipping)     |
| $colShippingPaid       | Shipping Paid                    | Amount paid by the customer for shipping         |
| $colTotalShippingPaid  | Total Shipping Paid              | Sum of shipping paid for all orders by recipient |
| $colShippingCost       | Shipping Cost                    | Actual shipping cost for the order               |
| $colRefundAmount       | Refund amount                    | Amount to refund to the customer                 |

---

## Code Walkthrough

### 1. Helper Functions

- **CleanCurrency($val):** Cleans up currency values by removing dollar signs and any non-numeric characters except `.` and `-`.
- **SafeDecimal($val):** Converts a cleaned currency string to a decimal number, defaulting to 0 if the value is empty or invalid.
- **ConvertTo-DataTable($Data):** Converts an array of PowerShell objects (from CSV) into a DataTable for use in the grid.
- **Add-ColumnsIfMissing($dt, $cols):** Ensures all required columns exist in the DataTable, adding any that are missing.
- **Remove-EmptyRows($dt):** Removes rows that are empty or missing a recipient.
- **Format-NumberWithCommas($num):** Formats numbers with commas for display in the UI.

### 2. Refund and Stats Calculation

- **Update-Refunds($dt):** Calculates shipping totals and refund amounts for each recipient. Groups rows by recipient, sums shipping paid and cost, and sets totals/refund amount in the first row for each recipient.
- **Add-TotalsRowAndFormat($grid, $dt, $boldFont):** Removes any existing Totals rows and adds a new Totals row with sums for each column. Formats the Totals row and headers.
- **Show-Stats($dt):** Calculates and displays summary statistics in the Stats tab (order count, refund rate, totals, averages, medians, etc.).
- **Show-CustomerStats($dt):** Calculates and displays customer statistics in the Customer Stats tab (total customers, average purchases, repeat rate, top customers, shipping stats).

### 3. UI Setup

- **Form and Controls:** Sets up the main window and controls (buttons, tabs, grids). Uses a top panel for buttons and a TabControl for Data, Stats, and Customer Stats tabs. Grids are only added to tabs after data is loaded.
- **DataGridView:** Displays the order data and statistics in table format. Formatting, row numbering, and colors are applied for clarity. Grids are hidden on startup and only shown after loading data.
- **Double Buffering:** Enabled for smoother grid scrolling.

### 4. Cell Formatting

- **Highlighting Logic:**
  - First row for each recipient: light blue
  - Missing/invalid shipping paid/cost: red
  - Refund amount: green (valid), yellow (zero/negative), red (missing data)
  - Totals row: bold

### 5. Button Actions

- **Load Purchases:** Loads order data from a CSV file. Updates grid, recalculates refunds, adds totals row, and enables export/recalc buttons.
- **Recalc:** Recalculates refunds and totals after data is edited. Updates stats tabs.
- **Export Purchases:** Saves the processed data to a new CSV file. Removes totals row before saving, restores after.
- **Export Stats:** Saves the statistics to a new CSV file.

### 6. Application Run and Cleanup

- **Form Events:** Handles layout, resizing, and cleanup on form closing. Ensures controls are only added once and layout logic is centralized.
- **Garbage Collection:** Cleans up grid data and triggers garbage collection on form close.

---

## Maintenance & Extension

- **Add Columns:** Update `$col...` variables and helper functions if your CSV changes.
- **Change Highlighting:** Edit cell formatting logic to adjust colors or rules.
- **Add Features:** Use the same button/event pattern for new actions.
- **Error Handling:** Add logging or status messages as needed for better feedback.
- **UI Customization:** Ensure each control is only added once and positioned as needed. Avoid duplicate additions or repeated code.

---

## Common Mistakes & How to Avoid Them

- **Column Name Mismatch:**
  - Always check that the script variables match your CSV column names exactly.
  - If you change your CSV headers, update the variables at the top of the script.
- **Running in VS Code or ISE:**
  - For best results, run the script in a standalone PowerShell window.
- **Editing Data:**
  - After editing data in the grid, always click "Recalc" to update calculations.
- **Saving Data:**
  - The app removes the totals row before saving to avoid issues in the output file.

---

## Redundancy Cleanup & Layout Logic

- The script now avoids duplicate UI control additions and repeated tab/grid initializations.
- Data and Stats grids are only added to their tabs after data is loaded, ensuring a clean UI on startup.
- Controls are only added to the form once, and layout logic is set in a single place for maintainability.
- Visibility logic for grids ensures users only see relevant data after loading.

---

## Contact & Support

For questions or improvements, contact the project owner or submit issues via the repository.

---

*This guide is intended for developers maintaining or extending the Shipping Refund Calculator. If you are new to PowerShell or Windows Forms, read each function and comment carefully, and test changes in a safe environment before deploying.*
