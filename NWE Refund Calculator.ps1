<#
=============================================================
PSRefundCalculator - NWE Refund Calculator
Copyright (c) 2025 palmersoftware

Licensed under the MIT License.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
=============================================================
NWE Refund Calculator - PowerShell GUI
Loads purchase data from CSV, calculates refunds, displays stats
Modern UI, robust error handling, and export features
=============================================================
#>

$grid = New-Object Windows.Forms.DataGridView # create main grid
<#
Input: Array of PowerShell objects (from CSV)
Output: DataTable for use in grid
#>
function ConvertTo-DataTable {
    param([Parameter(Mandatory)][object[]]$Data)
    $dt = New-Object System.Data.DataTable # create empty DataTable
    if ($Data.Count -eq 0) { return $dt } # return empty if no data
    $props = $Data[0].PSObject.Properties | ForEach-Object { $_.Name } # get property names
    foreach ($p in $props) { [void]$dt.Columns.Add($p) } # add columns
    foreach ($row in $Data) {
        $dr = $dt.NewRow() # create new row
        foreach ($p in $props) { $dr[$p] = $row.$p } # copy property value
        $dt.Rows.Add($dr) # add row
    }
    return $dt # return DataTable
}
<#
Safely parses a cleaned currency string to a decimal value.
Input: String with possible currency formatting
Output: Decimal value (0 if empty or invalid)
#>
function SafeDecimal($val) {
    $clean = CleanCurrency $val # remove currency formatting
    if ([string]::IsNullOrWhiteSpace($clean)) { return 0 } # treat empty as zero
    [decimal]::Parse($clean) # convert string to decimal
}
# -------------------------------
<#
Cleans currency strings by removing all non-numeric characters except '.' and '-'.
Input: String with possible currency formatting
Output: Cleaned numeric string
#>
function CleanCurrency($val) {
    if ($null -eq $val) { return "0" } # treat null as zero
    return ($val -replace '[^0-9.-]', '') # remove non-numeric chars
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

############################################################
# Define CSV column headers as variables
# Update these variables if your CSV uses different column names
############################################################
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
<#
Ensures all required columns exist in the DataTable, adding any that are missing.
Input: DataTable, array of required column names
Output: DataTable with all required columns
#>
function Add-ColumnsIfMissing($dt, [string[]]$cols) {
    # Loop through each required column name
    foreach ($col in $cols) {
        if (-not $dt.Columns.Contains($col)) { [void]$dt.Columns.Add($col) } # add missing column
    }
}

# -------------------------------
# Remove empty rows or rows without a recipient
<#
Removes rows that are empty or missing a recipient.
Input: DataTable
Output: DataTable with only valid rows
#>
function Remove-EmptyRows($dt) {
    # Loop through a copy of the rows (so we can safely remove rows)
    foreach ($row in @($dt.Rows)) {
        $isEmpty = $null -eq ($row.ItemArray | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ }) # check for empty row
        if ($isEmpty -or -not $row.$colRecipient) { $dt.Rows.Remove($row) } # remove if empty or missing recipient
    }
}

# -------------------------------
# Recalculate Shipping and Refunds
<#
Calculates shipping totals and refund amounts for each recipient.
Input: DataTable
Output: DataTable with calculated totals and refunds
#>
function Update-Refunds($dt) {
    # Make sure all required columns exist, and remove any empty rows
    Add-ColumnsIfMissing $dt $colShippingPaid,$colShippingCost,$colRefundAmount # ensure columns exist
    Remove-EmptyRows $dt # remove empty rows

    # Group rows by recipient so we can calculate totals per customer
    foreach ($group in ($dt | Group-Object $colRecipient)) {
        $totalPaid = 0; $totalCost = 0 # initialize totals
        foreach ($r in $group.Group) {
            $paid=0; [decimal]::TryParse((CleanCurrency $r.$colShippingPaid),[ref]$paid)|Out-Null; $totalPaid += $paid # sum shipping paid
            $cost=0; [decimal]::TryParse((CleanCurrency $r.$colShippingCost),[ref]$cost)|Out-Null; $totalCost += $cost # sum shipping cost
        }
        $first=$true # flag for first row
        foreach ($r in $group.Group) {
            if ($first) {
                $r.$colRefundAmount = if ($totalPaid -ne 0 -or $totalCost -ne 0){$totalPaid - $totalCost}else{''} # set refund
                $first=$false # only first row
            } else {
                $r.$colRefundAmount = '' # blank for others
            }
        }
    }
}

<#
Removes any existing Totals rows and adds a new Totals row with sums for each column.
Input: Grid, DataTable, Font for formatting
Output: Grid updated with formatted Totals row
#>
function Add-TotalsRowAndFormat {
    param($grid, $dt, $boldFont)
    # Remove any existing Totals rows to avoid duplicate totals
    for ($i = $dt.Rows.Count-1; $i -ge 0; $i--) {
        if ($dt.Rows[$i].$colOrder -eq "TOTAL") { $dt.Rows.RemoveAt($i) } # remove old totals
    }
    # Create a new Totals row and calculate sums for each relevant column
    $totals = $dt.NewRow() # create totals row
    if ($dt.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" } # set label
    $totals.$colQuantity = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colQuantity } | Measure-Object -Sum).Sum) # sum quantity
    $totals.$colOrderTotal = '{0:N2}' -f (($dt | ForEach-Object { SafeDecimal $_.$colOrderTotal } | Measure-Object -Sum).Sum) # sum order total
    $totals.$colShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingPaid}|Measure-Object -Sum).Sum) # sum shipping paid
    if ($dt.Columns.Contains($colTotalShippingPaid)) {
                $totals.$colTotalShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colTotalShippingPaid}|Measure-Object -Sum).Sum)
            }
    $totals.$colShippingCost = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colShippingCost}|Measure-Object -Sum).Sum) # sum shipping cost
    $totals.$colRefundAmount = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colRefundAmount}|Measure-Object -Sum).Sum) # sum refund
    # Remove all blank rows before adding Totals (cleanup)
    for ($i = $dt.Rows.Count-1; $i -ge 0; $i--) {
        if (($dt.Rows[$i].ItemArray -join '').Trim() -eq '') { $dt.Rows.RemoveAt($i) } # remove blank rows
    }
    $dt.Rows.Add($totals) # add totals row
    # Remove all blank rows after Totals (cleanup)
    for ($i = $dt.Rows.Count-1; $i -ge 0; $i--) {
        if (($dt.Rows[$i].ItemArray -join '').Trim() -eq '' -and $dt.Rows[$i].$colOrder -ne "TOTAL") { $dt.Rows.RemoveAt($i) } # remove blank after totals
    }
    # Set the grid's data source to the updated DataTable
    $grid.DataSource = $dt # update grid
    $grid.Refresh() # refresh grid
    $totRow = $grid.Rows.Count-1 # get totals row index
    $grid.Rows[$totRow].Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font = $boldFont } } # bold totals row
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold) # set header font
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb') # set header bg
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff') # set header fg
}
############################################################
$form = New-Object Windows.Forms.Form
$form.Text = "NWE Shipping Calculator" # set window title
$form.Size = '1400,800' # set window size
$form.StartPosition = "CenterScreen" # center window
$form.Icon = [System.Drawing.SystemIcons]::Application # set icon
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color
try {
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#273c75') # set foreground color
    $form.Paint.Add({
        $g = $_.Graphics # get graphics object
        $g.FillRectangle((New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#273c75'))), 0, 0, $form.Width, 32) # draw title bar
    })
} catch {} # ignore errors

# -------------------------------
############################################################
# Button setup and configuration
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

# Set button styles after creation
$btnLoad.FlatStyle = 'Flat'
$btnLoad.FlatAppearance.BorderSize = 0
$btnLoad.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
$btnLoad.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
$btnCalc.FlatStyle = 'Flat'
$btnCalc.FlatAppearance.BorderSize = 0
$btnCalc.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
$btnCalc.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
$btnExportPurchases.FlatStyle = 'Flat'
$btnExportPurchases.FlatAppearance.BorderSize = 0
$btnExportPurchases.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
$btnExportPurchases.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')
$btnExportStats.FlatStyle = 'Flat'
$btnExportStats.FlatAppearance.BorderSize = 0
$btnExportStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563eb')
$btnExportStats.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#fff')


############################################################
# Top panel setup for button layout
$topPanel = New-Object Windows.Forms.Panel # create top panel
$topPanel.Height = 44 # set height
$topPanel.Width = 600 # set width
$topPanel.Top = 8 # set top position
$topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2) # center horizontally
$topPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color
$btnLoad.Location = New-Object Drawing.Point(0,4) # set button position
$btnCalc.Location = New-Object Drawing.Point(150,4) # set button position
$btnExportPurchases.Location = New-Object Drawing.Point(300,4) # set button position
$btnExportStats.Location = New-Object Drawing.Point(450,4) # set button position
$topPanel.Controls.AddRange(@($btnLoad, $btnCalc, $btnExportPurchases, $btnExportStats)) # add buttons
$btnLoad.Cursor = [System.Windows.Forms.Cursors]::Hand # set cursor
$btnCalc.Enabled = $false # disable button
$btnExportPurchases.Enabled = $false # disable button
$btnExportStats.Enabled = $false # disable button

# -------------------------------
############################################################
# TabControl setup for main UI tabs
$tabControl = New-Object Windows.Forms.TabControl
$tabControl.Top = $topPanel.Bottom + 10
$tabControl.Left = 10
$tabControl.Width = 1360
$tabControl.Height = 700
$tabControl.Anchor = 'Top,Left,Right,Bottom'
$tabControl.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e')

    # Tab 1: Data
$tabData = New-Object Windows.Forms.TabPage # create data tab
$tabData.Text = 'Data' # set tab text
$tabData.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color

# Tab 2: Stats
$tabStats = New-Object Windows.Forms.TabPage # create stats tab
$tabStats.Text = 'Stats' # set tab text
$tabStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color
$gridStats.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set grid background

    # Tab 2: Customer Stats
$tabCustomerStats = New-Object Windows.Forms.TabPage # create customer stats tab
$tabCustomerStats.Text = 'Customer Stats' # set tab text
$tabCustomerStats.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#223a5e') # set background color
$gridCustomerStats = New-Object Windows.Forms.DataGridView # create customer stats grid
$gridCustomerStats.Dock = 'Fill' # dock fill
$gridCustomerStats.ReadOnly = $true # set read-only
$gridCustomerStats.AllowUserToAddRows = $false # disable add rows
$gridCustomerStats.ScrollBars = 'Both' # enable scrollbars
$gridCustomerStats.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#ececec') # set background color
$tabCustomerStats.Controls.Add($gridCustomerStats) # add grid to tab

$grid.Dock = [System.Windows.Forms.DockStyle]::Fill # dock fill
$grid.Visible = $false # ensure grid is hidden before adding to tab
$null # do not add grid to tab on startup
$null # do not add gridStats to tab on startup

$tabControl.TabPages.Clear() # clear tabs
$tabControl.TabPages.AddRange(@($tabData, $tabStats)) # add tabs

# -------------------------------
############################################################


    # Enable double buffering for smoother grid scrolling
$gridType = $grid.GetType()
$gridType = $grid.GetType() # get grid type
$gridStatsType = $gridStats.GetType() # get stats grid type
$doubleBufferedPropStats = $gridStatsType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance, NonPublic") # get double buffered property
$doubleBufferedPropStats.SetValue($gridStats, $true, $null) # enable double buffering
$doubleBufferedProp = $gridType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance, NonPublic") # get double buffered property
$doubleBufferedProp.SetValue($grid, $true, $null) # enable double buffering
$gridStatsType = $gridStats.GetType()
$doubleBufferedPropStats = $gridStatsType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance, NonPublic")
$doubleBufferedPropStats.SetValue($gridStats, $true, $null)
$doubleBufferedProp = $gridType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance, NonPublic")
$doubleBufferedProp.SetValue($grid, $true, $null)

    # Add row numbers to the row header using RowPostPaint event
 $grid.add_RowPostPaint({ # add row number event
    param($sender, $e)
    $rowIndex = $e.RowIndex + 2  # Start numbering at 2 for first data row (for clarity)
    $e.Graphics.DrawString(
        $rowIndex.ToString(),
        $e.InheritedRowStyle.Font,
        [System.Drawing.Brushes]::Black,
        $e.RowBounds.Location.X + 10,
        $e.RowBounds.Location.Y + 4
    )
})

# Add row numbers to gridStats row header
$gridStats.add_RowPostPaint({ # add row number event
    param($sender, $e)
    $rowIndex = $e.RowIndex + 2
    $e.Graphics.DrawString(
        $rowIndex.ToString(),
        $e.InheritedRowStyle.Font,
        [System.Drawing.Brushes]::Black,
        $e.RowBounds.Location.X + 10,
        $e.RowBounds.Location.Y + 4
    )
})

    # Set top-left header cell to '1' for clarity, similar to Excel
$grid.TopLeftHeaderCell.Value = '1' # set header cell value
$gridStats.TopLeftHeaderCell.Value = '1' # set header cell value

# -------------------------------
############################################################
# Layout: manually position controls for form and grid
$form.Add_Shown({ # on form shown
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.Refresh()
})
$form.Add_Resize({ # on form resize
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
})
    # Only add controls to the form once
$form.Controls.AddRange(@($topPanel, $tabControl)) # add controls to form



    # Position grid below topPanel and size to fill remaining space
    # Use Dock for layout; manual positioning removed

    # Bold font object for later use in grid formatting
$boldFont = New-Object System.Drawing.Font($grid.Font,[System.Drawing.FontStyle]::Bold) # create bold font

# -------------------------------
############################################################
# Initialize empty DataTable for grid


$grid.DataSource = $null # ensure grid has no data source on startup
$grid.AutoSizeColumnsMode = 'Fill' # auto size columns
$grid.BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml('#ececec') # set background color

    # Set formatting for stats grids
    # ...existing code...

############################################################
# Cell formatting event: highlights & bold
$grid.add_CellFormatting({ # cell formatting event
    param($src,$e)
    $row = $grid.Rows[$e.RowIndex]
    if (-not $row -or $row.IsNewRow) { return }
    $dt = $grid.DataSource
    # If Totals row, set all cells to bold font
    if ($row.Cells[$dt.Columns[$colOrder].Ordinal].Value -eq "TOTAL") {
        $row.Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font = $boldFont } }
        return
    }
    $colName = $dt.Columns[$e.ColumnIndex].ColumnName
    $currentRecipient = $row.Cells[$dt.Columns[$colRecipient].Ordinal].Value
    $isFirst = ($e.RowIndex -eq 0 -or $currentRecipient -ne $grid.Rows[$e.RowIndex-1].Cells[$dt.Columns[$colRecipient].Ordinal].Value)
    # First row per recipient gets light blue highlight for visibility
    if ($isFirst) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::LightSteelBlue }
    # Red highlight for missing Shipping Paid value
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
    # Refund Amount cell coloring and bold font
    elseif ($colName -eq $colRefundAmount -and $isFirst) {
        $row.Cells[$e.ColumnIndex].Style.Font=$boldFont
        $spVal=CleanCurrency $row.Cells[$dt.Columns[$colShippingPaid].Ordinal].Value
        $scVal=CleanCurrency $row.Cells[$dt.Columns[$colShippingCost].Ordinal].Value
        $refundVal=CleanCurrency $row.Cells[$e.ColumnIndex].Value
        $refund=0; $parsed=[decimal]::TryParse($refundVal,[ref]$refund)
        $hasRed=$false
        $spNum = 0 # initialize numeric value
        $spParsed = [decimal]::TryParse($spVal, [ref]$spNum) # try to parse string to decimal
        $scNum=0 # initialize numeric value
        $scParsed=[decimal]::TryParse($scVal,[ref]$scNum) # try to parse string to decimal
        if ([string]::IsNullOrWhiteSpace($spVal) -or -not $spParsed -or ($spNum -eq 0)) { $hasRed=$true }
        if ([string]::IsNullOrWhiteSpace($scVal) -or -not $scParsed) { $hasRed=$true } # check shipping cost
        if ($parsed -and $refund -le 0) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Yellow }
        elseif ($hasRed) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red } # highlight missing
        else { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Green } # highlight valid
    }
})

# ================================
<#
Loads purchase data from CSV, updates the grid, and recalculates stats.
If the user cancels, restores previous grid data seamlessly.
All UI and stats updates are handled here for consistency.
#>
function Load-Purchases {
    # Prepare UI and button states for loading purchases
    $btnLoad.FlatAppearance.BorderSize = 0
    Set-ButtonColors -active $true
    Set-ButtonCursors -active $true
    $tempData = $null
    # Store current grid data for restoration if needed
    if ($grid.DataSource -is [System.Data.DataTable] -and $grid.Rows.Count -gt 0) {
        $tempData = $grid.DataSource.Copy()
    }
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $grid.Visible = $false # Hide grid before loading
    # Open file dialog for CSV selection
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialogResult = $dlg.ShowDialog()
    if ($dialogResult -eq 'OK') {
        $csvPath = $dlg.FileName
        $csvData = Import-Csv $csvPath
        $dt = ConvertTo-DataTable -Data $csvData
        $global:dtOriginal = $dt.Copy()
        # Update grid with loaded data
        # Sort by recipient and order number for grouping and alphabetizing
        $dv = $dt.DefaultView
        $dv.Sort = "$colRecipient ASC, $colOrder ASC"
        $dtSorted = $dv.ToTable()
    if (-not $tabData.Controls.Contains($grid)) {
        $tabData.Controls.Add($grid) # add grid to tab only after loading
    }
    $grid.DataSource = $dtSorted
    $grid.Refresh()
    $grid.Visible = $true # Show grid after loading
        # Recalculate refunds and stats
        Update-Refunds $dtSorted
        Add-TotalsRowAndFormat $grid $dtSorted $boldFont
    if (-not $tabStats.Controls.Contains($gridStats)) {
        $tabStats.Controls.Add($gridStats) # add stats grid to tab only after loading
    }
    Show-Stats $dtSorted # Ensure Stats tab gridview is shown after loading data
    $btnCalc.Enabled = $true
    $btnExportPurchases.Enabled = $true
    $btnExportStats.Enabled = $true
    } else {
        # Restore previous grid data if cancelled
        if ($tempData) {
            $grid.DataSource = $tempData
            $grid.Refresh()
                $grid.Visible = $true
        }
    }
}

# Centralized button cursor logic
function Set-ButtonCursors {
    param([bool]$active)
    $cursor = if ($active) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
    $btnLoad.Cursor = $cursor
    $btnCalc.Cursor = $cursor
    $btnExportPurchases.Cursor = $cursor
    $btnExportStats.Cursor = $cursor
}

# Wire up the new modular function to the button
$btnLoad.Add_Click({ Load-Purchases })
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
            if ($dt.Columns.Contains($colTotalShippingPaid)) {
                $totals.$colTotalShippingPaid = '{0:N2}' -f (($dt|ForEach-Object{SafeDecimal $_.$colTotalShippingPaid}|Measure-Object -Sum).Sum)
            }
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
    # Show-Stats $dtStats # Remove duplicate call to prevent duplicated rows
    }
})

# Recalculate button
$btnCalc.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        # Use current grid data, including user edits
        $dtRaw = $grid.DataSource.Copy()
        # Remove totals row if present (to avoid double-counting)
        if ($dtRaw.Rows.Count -gt 0 -and $dtRaw.Rows[$dtRaw.Rows.Count-1].$colOrder -eq "TOTAL") {
            $dtRaw.Rows.RemoveAt($dtRaw.Rows.Count-1)
        }
        # Sort by recipient and order number for grouping and alphabetizing
        $dv = $dtRaw.DefaultView
        $dv.Sort = "$colRecipient ASC, $colOrder ASC"
        $dtSorted = $dv.ToTable()
        Update-Refunds $dtSorted
        Add-TotalsRowAndFormat $grid $dtSorted $boldFont
        # Update statistics tab with current grid data (excluding totals row)
        $dtStats = $dtSorted.Copy()
        if ($dtStats.Rows.Count -gt 0 -and $dtStats.Rows[$dtStats.Rows.Count-1].$colOrder -eq "TOTAL") {
            $dtStats.Rows.RemoveAt($dtStats.Rows.Count-1)
        }
    # Show-Stats $dtStats # Remove duplicate call to prevent duplicated rows
    }
})




# Function to calculate and display customer statistics
function Show-Stats {
    param($dt)
    # Always clear both rows and columns to prevent duplicate stats
    $gridStats.Rows.Clear()
    $gridStats.Columns.Clear()
    $gridStats.Columns.Add('Metric','Metric') | Out-Null
    $gridStats.Columns.Add('Value','Value') | Out-Null
    $gridStats.Columns[0].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $gridStats.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $gridStats.Columns[1].DefaultCellStyle.Format = 'N2'
    $gridStats.Visible = $true
    $gridStats.Refresh()
    if ($gridStats.Parent) { $gridStats.BringToFront() }
    $gridStats.Visible = $true
    if (-not $dt -or $dt.Rows.Count -eq 0) {
        $gridStats.Rows.Add('No data loaded','')
        return
    }
    # Filter out the Totals row so we only process actual data rows
    $realRows = $dt.Rows | Where-Object { $_.$colOrder -ne 'TOTAL' }

    # Extract and clean Order Total values for each row
    # - Remove any dollar signs or commas
    # - Convert valid strings to double
    # - Ignore any invalid or empty values
    $orderTotals = $realRows | ForEach-Object {
        $val = $_[$colOrderTotal] # Get the value from the current row
        if ($val -and ($val -is [string])) {
            $val = $val -replace '[\$,]', '' # Remove $ and ,
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null } # Convert to double if valid
        } elseif ($val -is [double]) { $val } else { $null } # Already a double
    } | Where-Object { $_ -ne $null } # Only keep valid numbers

    # Extract and clean Shipping Paid values for each row
    $shippingPaid = $realRows | ForEach-Object {
        $val = $_[$colShippingPaid]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', '' # Remove $ and ,
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }

    # Extract and clean Shipping Cost values for each row
    $shippingCosts = $realRows | ForEach-Object {
        $val = $_[$colShippingCost]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', '' # Remove $ and ,
            if ($val -match '^[\d\.]+$') { [double]$val } else { $null }
        } elseif ($val -is [double]) { $val } else { $null }
    } | Where-Object { $_ -ne $null }

    # Extract and clean Refund Amount values for each row
    $refunds = $realRows | ForEach-Object {
        $val = $_[$colRefundAmount]
        if ($val -and ($val -is [string]) ) {
            $val = $val -replace '[\$,]', '' # Remove $ and ,
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
    # Add rows to stats grid
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

# Calculates and displays customer statistics in the Customer Stats tab
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
    # Group by customer and calculate stats
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
    # Add rows to customer stats grid
    $gridCustomerStats.Rows.Add('Total customers', (Format-NumberWithCommas($totalCustomers)))
    $gridCustomerStats.Rows.Add('Average purchases per customer', $avgPurchases)
    $gridCustomerStats.Rows.Add('Median purchases per customer', $medPurchases)
    $gridCustomerStats.Rows.Add('Repeat purchase rate (%)', $repeatRate)
    foreach ($cust in $top5Customers) { $gridCustomerStats.Rows.Add("Top by purchase count", "$($cust.Name) ($($cust.Count))") }
    $gridCustomerStats.Rows.Add('Average shipping paid (what customer pays)', $avgShippingPaid)
    $gridCustomerStats.Rows.Add('Average shipping cost (what seller pays)', $avgShippingCost)
}

# Function to calculate and display purchase statistics


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

<#
Formats numbers with commas for display in the UI.
Input: Number (int or double)
Output: String with thousands separators
#>
function Format-NumberWithCommas($num) {
    if ($num -is [double] -or $num -is [int]) { return ($num -f "N0") } # format with commas
    else { return $num } # return as-is
}

<#
Handles cleanup when the form is closing.
Clears all grid data and triggers garbage collection.
#>
$form.add_FormClosing({ # on form closing
    # When the form is closing, clear all grid data and force garbage collection
    $grid.DataSource = $null
    $grid.Rows.Clear()
    $gridStats.DataSource = $null
    $gridStats.Rows.Clear()
    [System.GC]::Collect()
})

# Runs the main application form.
$form.ShowDialog() # run main form