# -------------------------------
# Converts an array of PSObjects (from Import-Csv) to a DataTable for DataGridView
function ConvertTo-DataTable {
    param([Parameter(Mandatory)][object[]]$Data)
    $dt = New-Object System.Data.DataTable
    if ($Data.Count -eq 0) { return $dt }
    $props = $Data[0].PSObject.Properties | ForEach-Object { $_.Name }
    foreach ($p in $props) { [void]$dt.Columns.Add($p) }
    foreach ($row in $Data) {
        $dr = $dt.NewRow()
        foreach ($p in $props) { $dr[$p] = $row.$p }
        $dt.Rows.Add($dr)
    }
    return $dt
}
# Safely parses a cleaned currency string to a decimal value.
# Returns 0 if the value is empty or invalid.
# -------------------------------
function SafeDecimal($val) {
    $clean = CleanCurrency $val
    if ([string]::IsNullOrWhiteSpace($clean)) { return 0 }
    [decimal]::Parse($clean)
}
# -------------------------------
# Cleans currency strings by removing all non-numeric characters except '.' and '-'.
# This allows users to enter values like "$12.34" and still have them parsed correctly.
# -------------------------------
function CleanCurrency($val) {
    if ($null -eq $val) { return "0" }
    return ($val -replace '[^0-9.-]', '')
}
# =========================================
# NWE Shipping Calculator - PowerShell GUI
# =========================================
# This application:
# - Loads CSV orders
# - Calculates shipping totals, refund amounts per recipient
# - Highlights missing values or issues
# - Displays totals and formatting for easy review
# =========================================

# Load Windows Forms library for GUI controls
Add-Type -AssemblyName System.Windows.Forms
# Load Drawing library for colors and fonts
Add-Type -AssemblyName System.Drawing

# -------------------------------
# Define CSV column headers as variables.
# If your CSV file uses different column names, update these variables to match.
# -------------------------------
$colOrder = "Order #"
$colItemName = "Item Name"
$colRecipient = "Recipient"
$colQuantity = "Quantity"
$colOrderTotal = "Order Total"
$colShippingPaid = "Shipping Paid"
$colTotalShippingPaid = "Total Shipping Paid"
$colShippingCost = "Shipping Cost"
$colRefundAmount = "Refund amount"

# -------------------------------
# Convert array of objects (CSV) to a DataTable
# -------------------------------
# Converts an array of PowerShell objects (from CSV) into a DataTable for use in the grid.
function ConvertTo-DataTable {
    param([Parameter(Mandatory)][object[]]$Data)
    $dt = New-Object System.Data.DataTable    # Create empty DataTable

    if (-not $Data -or $Data.Count -eq 0) { return $dt }  # Return empty if no data

    # Add a column for each property in the first object
    foreach ($prop in $Data[0].PSObject.Properties.Name) {
        [void]$dt.Columns.Add($prop)
    }

    # Fill each row with values from CSV objects
    foreach ($obj in $Data) {
        $row = $dt.NewRow()
        foreach ($prop in $obj.PSObject.Properties) {
            $row[$prop.Name] = $prop.Value
        }
        $dt.Rows.Add($row)
    }

    return ,$dt   # Return DataTable as array
}

# -------------------------------
# Ensure required columns exist in DataTable
# -------------------------------
# Ensures all required columns exist in the DataTable, adding any that are missing.
function Add-ColumnsIfMissing($dt, [string[]]$cols) {
    foreach ($col in $cols) {
        if (-not $dt.Columns.Contains($col)) {
            [void]$dt.Columns.Add($col)  # Add missing column
        }
    }
}

# -------------------------------
# Remove empty rows or rows without a recipient
# -------------------------------
# Removes rows that are empty or missing a recipient.
# This keeps the data clean and prevents calculation errors.
function Remove-EmptyRows($dt) {
    foreach ($row in @($dt.Rows)) {    # Loop through a copy to avoid modifying collection during iteration
        $isEmpty = $null -eq ($row.ItemArray | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ })
        if ($isEmpty -or -not $row.$colRecipient) {  # If row has no data or no recipient
            $dt.Rows.Remove($row)    # Remove row
        }
    }
}

# -------------------------------
# Recalculate Shipping and Refunds
# -------------------------------
# Calculates shipping totals and refund amounts for each recipient.
# Only the first row per recipient shows totals; others are left blank for clarity.
function Update-Refunds($dt) {
    Add-ColumnsIfMissing $dt $colShippingPaid,$colShippingCost,$colTotalShippingPaid,$colRefundAmount  # Ensure all needed columns exist
    Remove-EmptyRows $dt   # Remove empty rows

    # Move Total Shipping Paid column immediately after Shipping Paid for better readability
    $dt.Columns[$colTotalShippingPaid].SetOrdinal($dt.Columns[$colShippingPaid].Ordinal + 1)

    # Group rows by recipient so we can sum shipping paid/cost for each customer
    foreach ($group in ($dt | Group-Object $colRecipient)) {
        $totalPaid = 0; $totalCost = 0

        # Sum Shipping Paid and Shipping Cost for this recipient
        foreach ($r in $group.Group) {
            $paid=0; [decimal]::TryParse((CleanCurrency $r.$colShippingPaid),[ref]$paid)|Out-Null; $totalPaid += $paid
            $cost=0; [decimal]::TryParse((CleanCurrency $r.$colShippingCost),[ref]$cost)|Out-Null; $totalCost += $cost
        }

        $first=$true
        foreach ($r in $group.Group) {
            if ($first) {
                # Only first row per recipient gets totals and refund
                $r.$colTotalShippingPaid = if ($totalPaid -ne 0){$totalPaid}else{""}
                $r.$colRefundAmount = if ($totalPaid -ne 0 -or $totalCost -ne 0){$totalPaid - $totalCost}else{""}
                $first=$false
            } else {
                # Other rows for this recipient are left blank for totals/refund
                $r.$colTotalShippingPaid = ""; $r.$colRefundAmount = ""
            }
        }
    }
}

# -------------------------------
# GUI Form Setup
function Add-TotalsRowAndFormat {
    param($grid, $dt, $boldFont)
    # Remove all existing Totals rows
    for ($i = $dt.Rows.Count-1; $i -ge 0; $i--) {
        if ($dt.Rows[$i].$colOrder -eq "TOTAL") {
            $dt.Rows.RemoveAt($i)
        }
    }
    # Add new Totals row
    $totals = $dt.NewRow()
    if ($dt.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" }
    $totals.$colQuantity = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colQuantity } | Measure-Object -Sum).Sum)
    $totals.$colOrderTotal = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colOrderTotal } | Measure-Object -Sum).Sum)
    $totals.$colShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingPaid}|Measure-Object -Sum).Sum)
    $totals.$colTotalShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colTotalShippingPaid}|Measure-Object -Sum).Sum)
    $totals.$colShippingCost = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingCost}|Measure-Object -Sum).Sum)
    $totals.$colRefundAmount = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colRefundAmount}|Measure-Object -Sum).Sum)
    $dt.Rows.Add($totals)
    $grid.DataSource = $dt
    $grid.Refresh()
    $totRow = $grid.Rows.Count-1
    $grid.Rows[$totRow].Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font = $boldFont } }
}
# -------------------------------
$form = New-Object Windows.Forms.Form
$form.Text = "NWE Shipping Calculator"
$form.Size = '1400,800'
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Application
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # Medium dark blue
# Set window title bar color (where possible)
try {
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e')
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#273c75')
    $form.Paint.Add({
        $g = $_.Graphics
        $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#273c75'))), 0, 0, $form.Width, 32)
    })
} catch {}

# -------------------------------
# Buttons
$btnLoad.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLoad.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLoad.FlatStyle = 'Flat'
$btnLoad.FlatAppearance.BorderSize = 0
$btnLoad.Cursor = [System.Windows.Forms.Cursors]::Default
$btnCalc.Cursor = [System.Windows.Forms.Cursors]::Default
$btnExportPurchases.Cursor = [System.Windows.Forms.Cursors]::Default
$btnExportStats.Cursor = [System.Windows.Forms.Cursors]::Default
# -------------------------------

$btnLoad = New-Object Windows.Forms.Button; $btnLoad.Text="Load Purchases CSV"; $btnLoad.Size='140,36'
$btnCalc = New-Object Windows.Forms.Button; $btnCalc.Text="Recalc"; $btnCalc.Size='140,36'
$btnExportPurchases = New-Object Windows.Forms.Button; $btnExportPurchases.Text="Export Purchases"; $btnExportPurchases.Size='140,36'
$btnExportStats = New-Object Windows.Forms.Button; $btnExportStats.Text="Export Stats"; $btnExportStats.Size='140,36'


# Top panel to hold buttons neatly

# Use a panel for manual positioning, far upper left
$topPanel = New-Object Windows.Forms.Panel
$topPanel.Height = 44
$topPanel.Width = 600
$topPanel.Top = 8
$topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2) # Auto-center horizontally
$topPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # Medium dark blue
$btnLoad.Location = New-Object Drawing.Point(0,4)
$btnCalc.Location = New-Object Drawing.Point(150,4)
$btnExportPurchases.Location = New-Object Drawing.Point(300,4)
$btnExportStats.Location = New-Object Drawing.Point(450,4)
$topPanel.Controls.AddRange(@($btnLoad, $btnCalc, $btnExportPurchases, $btnExportStats))
$btnLoad.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCalc.Enabled = $false
$btnExportPurchases.Enabled = $false
$btnExportStats.Enabled = $false

# -------------------------------
# TabControl setup
# -------------------------------
$tabControl = New-Object Windows.Forms.TabControl
$tabControl.Top = $topPanel.Bottom + 10
$tabControl.Left = 10
$tabControl.Width = 1360
$tabControl.Height = 700
$tabControl.Anchor = 'Top,Left,Right,Bottom'
$tabControl.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # Medium dark blue

# Tab 1: Data
# Tab 1: Data
$tabData = New-Object Windows.Forms.TabPage
$tabData.Text = 'Data'
$tabData.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # Medium dark blue
# Only set properties after $grid is created

# Tab 2: Customer Stats
$tabCustomerStats = New-Object Windows.Forms.TabPage
$tabCustomerStats.Text = 'Customer Stats'
$tabCustomerStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e')
$gridCustomerStats = New-Object Windows.Forms.DataGridView
$gridCustomerStats.Dock = 'Fill'
$gridCustomerStats.ReadOnly = $true
$gridCustomerStats.AllowUserToAddRows = $false
$gridCustomerStats.ScrollBars = 'Both'
$gridCustomerStats.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#ececec')
$tabCustomerStats.Controls.Add($gridCustomerStats)

# Tab 3: Purchase Stats
 # Add tabs to TabControl
$tabPurchases = New-Object Windows.Forms.TabPage
$tabPurchases.Text = 'Purchases'
$tabPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e')
$tabPurchases.Controls.Add($grid)

 $tabStats = New-Object Windows.Forms.TabPage
 $tabStats.Text = 'Stats'
 $gridStats = New-Object Windows.Forms.DataGridView
 $gridStats.Dock = 'Fill'
 $gridStats.ReadOnly = $true
 $gridStats.AllowUserToAddRows = $false
 $gridStats.ScrollBars = 'Both'
 $tabStats.Controls.Add($gridStats)

 $tabControl.TabPages.Clear()
 $tabControl.TabPages.AddRange(@($tabPurchases, $tabStats))

# -------------------------------
# DataGridView setup
# -------------------------------

$grid = New-Object Windows.Forms.DataGridView
$grid.AutoGenerateColumns = $true
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $false
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font,[System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::EnableResizing
$grid.ColumnHeadersHeight = 30
$grid.RowsDefaultCellStyle.Font = $grid.Font
$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
$grid.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#f5f6fa')
$grid.GridColor = [System.Drawing.ColorTranslator]::FromHtml('#dcdde1')
$grid.BorderStyle = 'FixedSingle'
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb') # Modern blue accent
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
$grid.DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#f7f8fa') # Modern neutral
$grid.DefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222') # High contrast text
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml('#e0e7ff') # Subtle blue selection
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef') # Subtle contrast for rows
$btnLoad.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb') # Modern blue accent
$btnLoad.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
$btnCalc.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef') # Subtle panel color
$btnCalc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
$btnExportPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef') # Subtle panel color
$btnExportPurchases.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
$btnExportStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef')
$btnExportStats.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')

# Enable double buffering for smoother scrolling
$gridType = $grid.GetType()
$doubleBufferedProp = $gridType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance, NonPublic")
$doubleBufferedProp.SetValue($grid, $true, $null)

# Add row numbers to the row header using RowPostPaint event
$grid.add_RowPostPaint({
    param($sender, $e)
    $rowIndex = $e.RowIndex + 2  # Start numbering at 2 for first data row
    $e.Graphics.DrawString(
        $rowIndex.ToString(),
        $e.InheritedRowStyle.Font,
        [System.Drawing.Brushes]::Black,
        $e.RowBounds.Location.X + 10,
        $e.RowBounds.Location.Y + 4
    )
})

# Set top-left header cell to '1' for clarity, like Excel
$grid.TopLeftHeaderCell.Value = '1'

# -------------------------------
# Layout: manually position controls
$form.Add_Shown({
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.Refresh()
})
$form.Add_Resize({
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
})
# Only add controls once
$form.Controls.AddRange(@($topPanel, $tabControl))



# Position grid below topPanel and size to fill remaining space
# Only use Dock for layout; remove manual positioning

# Bold font object for later use
$boldFont = New-Object System.Drawing.Font($grid.Font,[System.Drawing.FontStyle]::Bold)

# -------------------------------
# Initialize empty DataTable for grid
# -------------------------------

$emptyTable = New-Object System.Data.DataTable
@($colOrder,$colItemName,$colRecipient,$colQuantity,$colOrderTotal,$colShippingPaid,$colTotalShippingPaid,$colShippingCost,$colRefundAmount) | ForEach-Object { [void]$emptyTable.Columns.Add($_) }
$grid.DataSource = $emptyTable
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#ececec') # Soft gray always

# Set formatting for stats grids
# ...existing code...

# ================================
# Cell formatting event: highlights & bold
# ================================
$grid.add_CellFormatting({
    param($src,$e)
    $row = $grid.Rows[$e.RowIndex]
    if (-not $row -or $row.IsNewRow) { return }
    $dt = $grid.DataSource
    # If Totals row, always set all cells to bold
    if ($row.Cells[$dt.Columns[$colOrder].Ordinal].Value -eq "TOTAL") {
        $row.Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font = $boldFont } }
        return
    }
    $colName = $dt.Columns[$e.ColumnIndex].ColumnName
    $currentRecipient = $row.Cells[$dt.Columns[$colRecipient].Ordinal].Value
    $isFirst = ($e.RowIndex -eq 0 -or $currentRecipient -ne $grid.Rows[$e.RowIndex-1].Cells[$dt.Columns[$colRecipient].Ordinal].Value)
    # First row per recipient gets light blue highlight
    if ($isFirst) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::LightSteelBlue }
    # Red highlight for missing Shipping Paid
    if ($colName -eq $colShippingPaid) {
        $val=CleanCurrency $row.Cells[$e.ColumnIndex].Value
        $num=0; $parsed=[decimal]::TryParse($val,[ref]$num)
        if ([string]::IsNullOrWhiteSpace($val) -or -not $parsed -or ($num -eq 0)) {
            $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red
        }
    }
    # Red highlight for missing Shipping Cost (only first row per recipient)
    elseif ($colName -eq $colShippingCost -and $isFirst) {
        $val=CleanCurrency $row.Cells[$e.ColumnIndex].Value
        $num=0; $parsed=[decimal]::TryParse($val,[ref]$num)
        if ([string]::IsNullOrWhiteSpace($val) -or -not $parsed) {
            $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red
        }
    }
    # Refund Amount coloring & bold
    elseif ($colName -eq $colRefundAmount -and $isFirst) {
        $row.Cells[$e.ColumnIndex].Style.Font=$boldFont
        $spVal=CleanCurrency $row.Cells[$dt.Columns[$colShippingPaid].Ordinal].Value
        $scVal=CleanCurrency $row.Cells[$dt.Columns[$colShippingCost].Ordinal].Value
        $refundVal=CleanCurrency $row.Cells[$e.ColumnIndex].Value
        $refund=0; $parsed=[decimal]::TryParse($refundVal,[ref]$refund)
        $hasRed=$false
        $spNum=0; $spParsed=[decimal]::TryParse($spVal,[ref]$spNum)
        $scNum=0; $scParsed=[decimal]::TryParse($scVal,[ref]$scNum)
        if ([string]::IsNullOrWhiteSpace($spVal) -or -not $spParsed -or ($spNum -eq 0)) { $hasRed=$true }
        if ([string]::IsNullOrWhiteSpace($scVal) -or -not $scParsed) { $hasRed=$true }
        if ($parsed -and $refund -le 0) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Yellow }
        elseif ($hasRed) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red }
        else { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Green }
    }
})

# ================================
# Button Actions
# ================================
# Load CSV
$btnLoad.Add_Click({
            $btnLoad.FlatAppearance.BorderSize = 0
            # Set all buttons to blue accent for consistency
            $btnLoad.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
            $btnCalc.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
            $btnExportPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
            $btnExportStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
            $btnLoad.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
            $btnCalc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
            $btnExportPurchases.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
            $btnExportStats.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
            # Set cursor to hand for enabled buttons
            $btnLoad.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnCalc.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnExportPurchases.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btnExportStats.Cursor = [System.Windows.Forms.Cursors]::Hand
    # Store current grid data in a temporary array
    $tempData = $null
    if ($grid.DataSource -is [System.Data.DataTable] -and $grid.Rows.Count -gt 0) {
        $tempData = $grid.DataSource.Copy()
    }
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $dialogResult = $dlg.ShowDialog()
        if ($dialogResult -eq 'OK') {
            $csvPath = $dlg.FileName
            $csvData = Import-Csv $csvPath
            $dt = ConvertTo-DataTable -Data $csvData
            $global:dtOriginal = $dt.Copy()
            # Sort/group by Recipient, then Order
            $dv = $dt.DefaultView
            $dv.Sort = "$colRecipient ASC, $colOrder ASC"
            $dtSorted = $dv.ToTable()
            Update-Refunds $dtSorted
            Add-TotalsRowAndFormat $grid $dtSorted $boldFont
            $grid.Dock = 'Fill'
            $grid.Visible = $true
            $tabPurchases.Controls.Clear()
            $tabPurchases.Controls.Add($grid)
            $tabControl.SelectedTab = $tabPurchases
            # Always pass a copy of the data table without the totals row to Show-Stats
            $dtStats = $global:dtOriginal.Copy()
            Show-Stats $dtStats
            $btnCalc.Enabled = $true
            $btnExportPurchases.Enabled = $true
            $btnExportStats.Enabled = $true
            # Set all buttons to blue accent for consistency ONLY after grid is populated
            if ($grid.DataSource -and $grid.Rows.Count -gt 0) {
                $btnLoad.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
                $btnCalc.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
                $btnExportPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
                $btnExportStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
                $btnLoad.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
                $btnCalc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
                $btnExportPurchases.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
                $btnExportStats.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
            }
        } elseif ($tempData) {
            # Restore previous grid data if user cancels
            $global:dtOriginal = $tempData.Copy()
            # Sort/group by Recipient, then Order
            $dv = $tempData.DefaultView
            $dv.Sort = "$colRecipient ASC, $colOrder ASC"
            $dtSorted = $dv.ToTable()
            Update-Refunds $dtSorted
            Add-TotalsRowAndFormat $grid $dtSorted $boldFont
            $grid.Dock = 'Fill'
            $grid.Visible = $true
            $tabPurchases.Controls.Clear()
            $tabPurchases.Controls.Add($grid)
            $tabControl.SelectedTab = $tabPurchases
            $dtStats = $global:dtOriginal.Copy()
            Show-Stats $dtStats
            $btnCalc.Enabled = $true
            $btnExportPurchases.Enabled = $true
            $btnExportStats.Enabled = $true
            # Restore default button colors (do not turn blue)
            $btnCalc.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef')
            $btnExportPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef')
            $btnExportStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#e9ecef')
            $btnCalc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
            $btnExportPurchases.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
            $btnExportStats.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#222')
        } elseif ($tempData) {
            # Restore previous grid data if user cancels
            $global:dtOriginal = $tempData.Copy()
            # Sort/group by Recipient, then Order
            $dv = $tempData.DefaultView
            $dv.Sort = "$colRecipient ASC, $colOrder ASC"
            $dtSorted = $dv.ToTable()
            Update-Refunds $dtSorted
            # Remove all existing totals rows before adding a new one
            for ($i = $dtSorted.Rows.Count-1; $i -ge 0; $i--) {
                if ($dtSorted.Rows[$i].$colOrder -eq "TOTAL") {
                    $dtSorted.Rows.RemoveAt($i)
                }
            }
            # Add totals row for display only (at end)
            $totals = $dtSorted.NewRow()
            if ($dtSorted.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" }
            $totals.$colQuantity = ($dtSorted | ForEach-Object { SafeDecimal $_.$colQuantity } | Measure-Object -Sum).Sum
            $totals.$colOrderTotal = ($dtSorted | ForEach-Object { SafeDecimal $_.$colOrderTotal } | Measure-Object -Sum).Sum
            $totals.$colShippingPaid=($dtSorted|ForEach-Object{SafeDecimal $_.$colShippingPaid}|Measure-Object -Sum).Sum
            $totals.$colTotalShippingPaid=($dtSorted|ForEach-Object{SafeDecimal $_.$colTotalShippingPaid}|Measure-Object -Sum).Sum
            $totals.$colShippingCost=($dtSorted|ForEach-Object{SafeDecimal $_.$colShippingCost}|Measure-Object -Sum).Sum
            $totals.$colRefundAmount=($dtSorted|ForEach-Object{SafeDecimal $_.$colRefundAmount}|Measure-Object -Sum).Sum
            $dtSorted.Rows.Add($totals)
            $grid.DataSource = $dtSorted
            $grid.Dock = 'Fill'
            $grid.Visible = $true
            $grid.Refresh()
            $tabPurchases.Controls.Clear()
            $tabPurchases.Controls.Add($grid)
            $tabControl.SelectedTab = $tabPurchases
            $dtStats = $global:dtOriginal.Copy()
            Show-Stats $dtStats
            $btnCalc.Enabled = $true
            $btnExportPurchases.Enabled = $true
            $btnExportStats.Enabled = $true
        }
})
$btnExportPurchases.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        # Workaround: temporarily store grid data, clear grid, run dialog, restore grid
        $tempData = $grid.DataSource.Copy()
        $grid.DataSource = $null
        $grid.Rows.Clear()
        $dt = $tempData
        $csv = $dt | ConvertTo-Csv -NoTypeInformation
        $dateStr = (Get-Date -Format 'yyyyMMdd')
        $fileName = "purchases_${dateStr}.csv"
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.FileName = $fileName
        $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq 'OK') { [System.IO.File]::WriteAllLines($dlg.FileName, $csv) }
        # Restore grid data and recalc
        $grid.DataSource = $tempData
        $grid.Refresh()
        if ($global:dtOriginal -is [System.Data.DataTable]) {
            $dtRaw = $grid.DataSource.Copy()
            # Remove all existing Totals rows before recalculating
            for ($i = $dtRaw.Rows.Count-1; $i -ge 0; $i--) {
                if ($dtRaw.Rows[$i].$colOrder -eq "TOTAL") {
                    $dtRaw.Rows.RemoveAt($i)
                }
            }
            $dv = $dtRaw.DefaultView
            $dv.Sort = "$colRecipient ASC, $colOrder ASC"
            $dt = $dv.ToTable()
            Update-Refunds $dt
            $totals = $dt.NewRow()
            if ($dt.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" }
            $totals.$colQuantity = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colQuantity } | Measure-Object -Sum).Sum)
            $totals.$colOrderTotal = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colOrderTotal } | Measure-Object -Sum).Sum)
            $totals.$colShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingPaid}|Measure-Object -Sum).Sum)
            $totals.$colTotalShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colTotalShippingPaid}|Measure-Object -Sum).Sum)
            $totals.$colShippingCost = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingCost}|Measure-Object -Sum).Sum)
            $totals.$colRefundAmount = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colRefundAmount}|Measure-Object -Sum).Sum)
            $dt.Rows.Add($totals)
            $grid.DataSource = $dt
            $grid.Refresh()
            $dtStats = $dt.Copy()
            for ($i = $dtStats.Rows.Count-1; $i -ge 0; $i--) {
                if ($dtStats.Rows[$i].$colOrder -eq "TOTAL") {
                    $dtStats.Rows.RemoveAt($i)
                }
            }
            Show-Stats $dtStats
        }
    }
})
$btnExportStats.Add_Click({
    # Workaround: temporarily store grid data, clear grid, run dialog, restore grid
    $tempData = $grid.DataSource.Copy()
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $csv = @('Metric,Value')
    foreach ($row in $gridStats.Rows) {
        if ($row.Cells.Count -eq 2) {
            $metric = $row.Cells[0].Value
            $value = $row.Cells[1].Value
            if ($metric -and $value) { $csv += "$metric,$value" }
        }
    }
    $dateStr = (Get-Date -Format 'yyyyMMdd')
    $fileName = "stats_${dateStr}.csv"
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.FileName = $fileName
    $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq 'OK') { [System.IO.File]::WriteAllLines($dlg.FileName, $csv) }
    # Restore grid data and recalc
    $grid.DataSource = $tempData
    $grid.Refresh()
    if ($global:dtOriginal -is [System.Data.DataTable]) {
        $dtRaw = $grid.DataSource.Copy()
    $dv = $dtRaw.DefaultView
    $dv.Sort = "$colRecipient ASC, $colOrder ASC"
    $dt = $dv.ToTable()
    Update-Refunds $dt
    Add-TotalsRowAndFormat $grid $dt $boldFont
        $dtStats = $dt.Copy()
        for ($i = $dtStats.Rows.Count-1; $i -ge 0; $i--) {
            if ($dtStats.Rows[$i].$colOrder -eq "TOTAL") {
                $dtStats.Rows.RemoveAt($i)
            }
        }
        Show-Stats $dtStats
    }
})

# Recalculate button
$btnCalc.Add_Click({
    if ($global:dtOriginal -is [System.Data.DataTable]) {
        # Use current grid data, including user edits
        $dtRaw = $grid.DataSource.Copy()
        # Remove totals row if present (to avoid double-counting)
        if ($dtRaw.Rows.Count -gt 0 -and $dtRaw.Rows[$dtRaw.Rows.Count-1].$colOrder -eq "TOTAL") {
            $dtRaw.Rows.RemoveAt($dtRaw.Rows.Count-1)
        }
        # --- Restore correct order: sort by Recipient, then Order # ---
        $dv = $dtRaw.DefaultView
        $dv.Sort = "$colRecipient ASC, $colOrder ASC"
        $dt = $dv.ToTable()
        Update-Refunds $dt
    Add-TotalsRowAndFormat $grid $dt $boldFont
        # Update statistics tab with current grid data (excluding totals row)
        $dtStats = $dt.Copy()
        if ($dtStats.Rows.Count -gt 0 -and $dtStats.Rows[$dtStats.Rows.Count-1].$colOrder -eq "TOTAL") {
            $dtStats.Rows.RemoveAt($dtStats.Rows.Count-1)
        }
        Show-Stats $dtStats
    }
})




# Function to calculate and display customer statistics
function Show-Stats {
    param($dt)
    $gridStats.Rows.Clear()
    $gridStats.Columns.Clear()
    $gridStats.Columns.Add('Metric','Metric') | Out-Null
    $gridStats.Columns.Add('Value','Value') | Out-Null
    $gridStats.Columns[0].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $gridStats.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $gridStats.Columns[1].DefaultCellStyle.Format = 'N2'
    if (-not $dt -or $dt.Rows.Count -eq 0) {
        $gridStats.Rows.Add('No data loaded','')
        return
    }
    $realRows = $dt.Rows | Where-Object { $_.$colOrder -ne 'TOTAL' }
    $orderTotals = $realRows | ForEach-Object {
        $val = $_[$colOrderTotal]
        if ($val -and ($val -is [string])) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $shippingPaid = $realRows | ForEach-Object {
        $val = $_[$colShippingPaid]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $shippingCosts = $realRows | ForEach-Object {
        $val = $_[$colShippingCost]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $refunds = $realRows | ForEach-Object {
        $val = $_[$colRefundAmount]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $numOrders = $orderTotals.Count
    $numRefunded = ($refunds | Where-Object { $_ -gt 0 }).Count
    $refundRate = if ($numOrders -gt 0) { [math]::Round(($numRefunded / $numOrders) * 100,2) } else { 0 }
    $avgOrder = if ($orderTotals.Count -gt 0) { [math]::Round(($orderTotals | Measure-Object -Average).Average,2) } else { 0 }
    $sortedOrders = $orderTotals | Sort-Object
    $medOrder = if ($sortedOrders.Count -gt 0) { $sortedOrders[[int](($sortedOrders.Count-1)/2)] } else { 0 }
    $maxOrder = if ($sortedOrders.Count -gt 0) { $sortedOrders[-1] } else { 0 }
    $minOrder = if ($sortedOrders.Count -gt 0) { $sortedOrders[0] } else { 0 }
    $avgShip = if ($shippingCosts.Count -gt 0) { [math]::Round(($shippingCosts | Measure-Object -Average).Average,2) } else { 0 }
    $sortedShip = $shippingCosts | Sort-Object
    $medShip = if ($sortedShip.Count -gt 0) { $sortedShip[[int](($sortedShip.Count-1)/2)] } else { 0 }
    $maxShip = if ($sortedShip.Count -gt 0) { $sortedShip[-1] } else { 0 }
    $minShip = if ($sortedShip.Count -gt 0) { $sortedShip[0] } else { 0 }
    $totalSales = ($orderTotals | Measure-Object -Sum).Sum
    $totalShippingPaid = ($shippingPaid | Measure-Object -Sum).Sum
    $totalShippingCost = ($shippingCosts | Measure-Object -Sum).Sum
    $totalRefunds = ($refunds | Measure-Object -Sum).Sum
    $gridStats.Rows.Add("Number of orders (Count of all orders)", (Format-NumberWithCommas($numOrders)))
    $gridStats.Rows.Add("Number of refunded orders (Count of orders with RefundAmount > 0)", (Format-NumberWithCommas($numRefunded)))
    $gridStats.Rows.Add("Refund rate (%) (Number of refunded orders / Number of orders * 100)", $refundRate)
    $gridStats.Rows.Add("Total sales (Sum of Order Total for all orders)","$" + (Format-NumberWithCommas($totalSales)))
    $gridStats.Rows.Add("Total shipping paid (Sum of Shipping Paid for all orders)","$" + (Format-NumberWithCommas($totalShippingPaid)))
    $gridStats.Rows.Add("Total shipping cost (Sum of Shipping Cost for all orders)","$" + (Format-NumberWithCommas($totalShippingCost)))
    $gridStats.Rows.Add("Total refunds (Sum of Refund Amount for all orders)","$" + (Format-NumberWithCommas($totalRefunds)))
    $gridStats.Rows.Add("Average order total (Total sales / Number of orders)","$" + (Format-NumberWithCommas($avgOrder)))
    $gridStats.Rows.Add("Median order total (Middle value of sorted Order Totals)","$" + (Format-NumberWithCommas($medOrder)))
    $gridStats.Rows.Add("Highest order total (Maximum Order Total)","$" + (Format-NumberWithCommas($maxOrder)))
    $gridStats.Rows.Add("Lowest order total (Minimum Order Total)","$" + (Format-NumberWithCommas($minOrder)))
    $gridStats.Rows.Add("Average shipping cost (Total shipping cost / Number of orders)","$" + (Format-NumberWithCommas($avgShip)))
    $gridStats.Rows.Add("Median shipping cost (Middle value of sorted Shipping Costs)","$" + (Format-NumberWithCommas($medShip)))
    $gridStats.Rows.Add("Highest shipping cost (Maximum Shipping Cost)","$" + (Format-NumberWithCommas($maxShip)))
    $gridStats.Rows.Add("Lowest shipping cost (Minimum Shipping Cost)","$" + (Format-NumberWithCommas($minShip)))
}
function Show-CustomerStats {
    param($dt)
    $gridCustomerStats.Rows.Clear()
    $gridCustomerStats.Columns.Clear()
    $gridCustomerStats.Columns.Add('Metric','Metric') | Out-Null
    $gridCustomerStats.Columns.Add('Value','Value') | Out-Null
    $gridCustomerStats.Columns[0].AutoSizeMode = 'AllCells'
    $gridCustomerStats.Columns[1].AutoSizeMode = 'Fill'
    $gridCustomerStats.Columns[1].DefaultCellStyle.Format = 'N2'
    if (-not $dt -or $dt.Rows.Count -eq 0) {
        $gridCustomerStats.Rows.Add('No data loaded','')
        return
    }
    $customers = $dt.Rows | ForEach-Object { $_[$colRecipient] } | Where-Object { $_ }
    $customerGroups = $customers | Group-Object
    $totalCustomers = $customerGroups.Count
    $purchasesPerCustomer = $customerGroups | ForEach-Object { $_.Count }
    $avgPurchases = [math]::Round(($purchasesPerCustomer | Measure-Object -Average).Average,2)
    $sortedPurchases = $purchasesPerCustomer | Sort-Object
    $medPurchases = if ($sortedPurchases.Count -gt 0) { $sortedPurchases[[int](($sortedPurchases.Count-1)/2)] } else { 0 }
    $repeatCustomers = $customerGroups | Where-Object { $_.Count -gt 1 }
    $repeatRate = if ($totalCustomers -gt 0) { [math]::Round(($repeatCustomers.Count / $totalCustomers) * 100,2) } else { 0 }
    $top5Customers = $customerGroups | Sort-Object Count -Descending | Select-Object -First 5
    # Shipping Paid and Shipping Cost stats
    $shippingPaid = $dt.Rows | ForEach-Object {
        $val = $_[$colShippingPaid]
        if ($val -and ($val -is [string])) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $shippingCosts = $dt.Rows | ForEach-Object {
        $val = $_[$colShippingCost]
        if ($val -and ($val -is [string])) {
            $val = $val -replace '[\$,]', ''
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }
    $avgShippingPaid = if ($shippingPaid.Count -gt 0) { [math]::Round(($shippingPaid | Measure-Object -Average).Average,2) } else { 0 }
    $sortedShippingPaid = $shippingPaid | Sort-Object
    $medShippingPaid = if ($sortedShippingPaid.Count -gt 0) { $sortedShippingPaid[[int](($sortedShippingPaid.Count-1)/2)] } else { 0 }
    $avgShippingCost = if ($shippingCosts.Count -gt 0) { [math]::Round(($shippingCosts | Measure-Object -Average).Average,2) } else { 0 }
    $sortedShippingCost = $shippingCosts | Sort-Object
    $medShippingCost = if ($sortedShippingCost.Count -gt 0) { $sortedShippingCost[[int](($sortedShippingCost.Count-1)/2)] } else { 0 }
    $gridCustomerStats.Rows.Add('Total customers', (Format-NumberWithCommas($totalCustomers)))
    $gridCustomerStats.Rows.Add('Average purchases per customer', $avgPurchases)
    $gridCustomerStats.Rows.Add('Median purchases per customer', $medPurchases)
    $gridCustomerStats.Rows.Add('Repeat purchase rate (%)', $repeatRate)
    foreach ($cust in $top5Customers) { $gridCustomerStats.Rows.Add("Top by purchase count", "$($cust.Name) ($($cust.Count))") }
    $gridCustomerStats.Rows.Add('Average shipping paid (what customer pays)', $avgShippingPaid)
    $gridCustomerStats.Rows.Add('Average shipping cost (what seller pays)', $avgShippingCost)
}

# Function to calculate and display purchase statistics
function Show-PurchaseStats {
    # Removed Show-PurchaseStats function and replaced with Show-Stats
    Show-Stats $dt
}

# Update Recalc button to refresh statistics
$btnCalc.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        $dt = $grid.DataSource
        Show-CustomerStats $dt
        Show-PurchaseStats $dt
    }
})

# Also refresh stats after loading CSV
$btnLoad.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        $dt = $grid.DataSource
        Show-CustomerStats $dt
        Show-PurchaseStats $dt
    }
})

# Utility function for formatting numbers with commas
function Format-NumberWithCommas($num) {
    if ($num -is [double] -or $num -is [int]) {
        return ($num -f "N0")
    } else {
        return $num
    }
}

# -------------------------------
# Form Closing event: cleanup
# -------------------------------
$form.add_FormClosing({
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $gridStats.DataSource = $null
    $gridStats.Rows.Clear()
    [System.GC]::Collect()
})

# ================================
# Run the form
# ================================
$form.ShowDialog()
