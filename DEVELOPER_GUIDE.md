# NWE Shipping Refund Calculator - Developer Guide

This guide provides a deep technical overview of the PowerShell GUI application, including code structure, logic, and maintenance tips. It is written for junior developers and explains every major part of the code and why it exists.

---

## Application Overview

- Loads order data from CSV
- Calculates shipping totals and refunds per recipient
- Highlights missing/invalid data
- Allows editing and saving results
- Built with Windows Forms for a simple GUI

---

## CSV Column Mapping

The script uses variables to reference CSV column names. This makes it easy to update the script if your CSV headers change. **You must ensure these variables match your CSV file's column names exactly.**

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

#### CleanCurrency($val)
- **Purpose:** Cleans up currency values by removing dollar signs and any non-numeric characters except `.` and `-`.
- **Why:** Users may enter values like "$12.34". This function ensures calculations work regardless of formatting.
- **How:** Uses a regular expression to strip unwanted characters.

#### SafeDecimal($val)
- **Purpose:** Converts a cleaned currency string to a decimal number, defaulting to 0 if the value is empty or invalid.
- **Why:** Prevents errors when parsing blank or malformed data.
- **How:** Calls `CleanCurrency`, checks for empty, then parses as decimal.

#### ConvertTo-DataTable($Data)
- **Purpose:** Converts an array of PowerShell objects (from CSV) into a DataTable for use in the grid.
- **Why:** DataTable is required for the DataGridView control to display and edit tabular data.
- **How:** Adds columns based on the first object's properties, then fills rows with values.

#### Add-ColumnsIfMissing($dt, $cols)
- **Purpose:** Ensures all required columns exist in the DataTable, adding any that are missing.
- **Why:** Prevents errors if the CSV is missing expected columns.
- **How:** Loops through the list of required columns and adds any that are missing.

#### Remove-EmptyRows($dt)
- **Purpose:** Removes rows that are empty or missing a recipient.
- **Why:** Keeps the data clean and prevents calculation errors.
- **How:** Checks each row for empty values or missing recipient and removes it.

#### Update-Refunds($dt)
- **Purpose:** Calculates shipping totals and refund amounts for each recipient.
- **Why:** Automates the refund calculation logic for all orders.
- **How:**
  - Groups rows by recipient
  - Sums shipping paid and cost for each group
  - Sets totals and refund amount in the first row for each recipient
  - Leaves other rows blank for clarity

### 2. UI Setup

#### Form and Controls
- **Purpose:** Sets up the main window and controls (buttons, grid).
- **Why:** Provides a user-friendly interface for loading, editing, and saving data.
- **How:**
  - Creates a Windows Form
  - Adds three buttons: Load CSV, Recalc, Save CSV
  - Uses a FlowLayoutPanel to arrange buttons
  - Sets up a DataGridView for displaying and editing data
  - Applies bold font for headers and totals

#### DataGridView
- **Purpose:** Displays the order data in a table format.
- **Why:** Allows users to view and edit data easily.
- **How:**
  - Initializes with all required columns
  - Sets properties for auto-sizing, editing, and appearance

### 3. Cell Formatting

#### Highlighting Logic
- **Purpose:** Visually marks important cells for easy review.
- **Why:** Helps users spot missing or invalid data quickly.
- **How:**
  - First row for each recipient: light blue
  - Missing/invalid shipping paid/cost: red
  - Refund amount: green (valid), yellow (zero/negative), red (missing data)
  - Totals row: bold

### 4. Button Actions

#### Load CSV
- **Purpose:** Loads order data from a CSV file.
- **Why:** Allows users to import their order data for processing.
- **How:**
  - Uses a single, reusable OpenFileDialog instance
  - Loads and processes data, recalculates refunds, adds totals row
  - Refreshes the grid

#### Recalc
- **Purpose:** Recalculates refunds and totals after data is edited.
- **Why:** Ensures calculations are up to date after changes.
- **How:**
  - Removes old totals row
  - Sorts data by recipient and order number
  - Recalculates refunds and adds a new totals row
  - Refreshes the grid

#### Save CSV
- **Purpose:** Saves the processed data to a new CSV file.
- **Why:** Allows users to export results for record-keeping or further use.
- **How:**
  - Removes totals row before saving
  - Uses a single, reusable SaveFileDialog instance
  - Converts DataTable to array of objects for Export-Csv
  - Restores totals row after saving
  - Silent save (no popups)

### 5. Application Run

- **Purpose:** Keeps the app open until the user closes the window.
- **Why:** Ensures the user can interact with the app as long as needed.
- **How:** Uses `[System.Windows.Forms.Application]::Run($form)` at the end of the script.

---

## Maintenance & Extension

- **Add Columns:** Update `$col...` variables and helper functions if your CSV changes.
- **Change Highlighting:** Edit `$grid.add_CellFormatting` logic to adjust colors or rules.
- **Add Features:** Use the same button/event pattern for new actions.
- **Error Handling:** Add logging or status messages as needed for better feedback.

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

## Extending the App

- Add new columns by updating the `$col...` variables and helper functions.
- Add new buttons or features by following the existing event handler pattern.
- Improve error handling by adding status messages or logging.
- For advanced features (Excel/PDF export, undo, etc.), consider using additional PowerShell modules or .NET libraries.

---

## Contact & Support

For questions or improvements, contact the project owner or submit issues via the repository.

---

*This guide is intended for developers maintaining or extending the Shipping Refund Calculator. If you are new to PowerShell or Windows Forms, read each function and comment carefully, and test changes in a safe environment before deploying.*

## Redundancy Cleanup & Layout Logic

- The script now avoids duplicate UI control additions and repeated panel positioning.
- Controls are only added to the form once, and layout logic is set in a single place for maintainability.
- If you extend or customize the UI, always check for duplicate additions or unnecessary repeated code.
