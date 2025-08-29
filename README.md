# PSRefundCalculator
To calculate owed refund to customers on PalmStreet. Import CSV to sum customers orders shipping paid and deduct from their order shipping cost to see amounts to refund.

Ensure updating the .csv's column header names in the .ps1 file for setting the variable names correctly

Example  
$colOrder = "Order #"  
$colItemName = "Item Name"  
$colRecipient = "Recipient"  
$colQuantity = "Quantity"  
$colOrderTotal = "Order Total"  
$colShippingPaid = "Shipping Paid"  
$colTotalShippingPaid = "Total Shipping Paid"  
$colShippingCost = "Shipping Cost"  
$colRefundAmount = "Refund amount"  
