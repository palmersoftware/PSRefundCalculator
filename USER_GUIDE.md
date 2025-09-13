# NWE Shipping Refund Calculator - User Guide

Welcome! This guide explains how to use the Shipping Refund Calculator application.

---

## What does this app do?

- Loads your order data from a CSV file
- Calculates shipping refunds for each recipient
- Highlights missing or invalid data for easy review
- Displays summary statistics and customer stats
- Lets you edit and save results to a new CSV file
- Exports purchases and stats to CSV

---

## CSV Format Requirements

Your CSV file should have the following columns. If your column names are different, update the script variables at the top of the PowerShell file to match:

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

> **Important:** If your CSV uses different column names, you must update the script variables at the top of the PowerShell file to match your file headers exactly. Otherwise, the app will not work correctly.

---

## How to Use

1. **Open the Application**
   - Run the PowerShell script in a regular PowerShell window.

2. **Load Your Orders**
   - Click the "Load Purchases CSV" button and select your CSV file.
   - The Data and Stats tables will appear only after loading data.

3. **Review and Edit Data**
   - Missing or invalid shipping info is highlighted in red.
   - Refunds are calculated automatically and shown in the Data table.
   - Totals are shown at the bottom of the Data table.
   - The Stats tab shows summary statistics; the Customer Stats tab shows customer-specific statistics.
   - You can edit any cell directly in the Data table. After editing, click "Recalc" to update totals and refunds.

4. **Export Results**
   - Click "Export Purchases" to save updated data to a new CSV file.
   - Click "Export Stats" to save statistics to a new CSV file. The app will overwrite files silently if you choose an existing filename.

---

## Application Layout Notes

- Data and Stats tables are hidden on startup and only shown after loading data.
- Controls are only added once for a clean UI.
- The top panel and tab controls are added to the form a single time, and their positions are set for clarity and maintainability.

---

## Tips

- If you see "Working On It..." when loading files, try running the app in a regular PowerShell window (not inside VS Code).
- Dollar signs and other symbols in shipping fields are OK—the app will ignore them for calculations.
- If you have questions or need help, contact https://github.com/palmersoftware/PSRefundCalculator

---

Thank you for using the Shipping Refund Calculator!
