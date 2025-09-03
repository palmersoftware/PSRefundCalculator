# NWE Shipping Refund Calculator - User Guide

Welcome! This guide explains how to use the Shipping Refund Calculator application.

---

## What does this app do?

- Loads your order data from a CSV file
- Calculates shipping refunds for each recipient
- Highlights missing or invalid data for easy review
- Lets you save the results to a new CSV file

---

## Table Columns (CSV Format)

Your CSV file should have the following columns, and the script variables must reference the exact column names as they appear in your file. If your column names are different, you must update the script variables to match:

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
   - Click the "Load CSV" button.
   - Select your CSV file containing order data.
   - Your data will appear in the table.

3. **Review Your Data**
   - The app highlights missing or invalid shipping info in red.
   - Refunds are calculated automatically and shown in the table.
   - Totals are shown at the bottom.

4. **Edit if Needed**
   - You can click and edit any cell directly in the table.
   - If you make changes, click "Recalc" to update totals and refunds.

5. **Save Your Results**
   - Click "Save CSV" to export the updated data to a new file.
   - The app will overwrite files silently if you choose an existing filename.

---

## Tips

- If you see "Working On It..." when loading files, try running the app in a regular PowerShell window (not inside VS Code).
- Dollar signs and other symbols in shipping fields are OK—the app will ignore them for calculations.
- If you have questions or need help, contact https://github.com/palmersoftware/PSRefundCalculator

---

Thank you for using the Shipping Refund Calculator!
