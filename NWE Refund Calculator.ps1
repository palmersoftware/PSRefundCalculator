# =========================================
# NWE Shipping Calculator - PowerShell GUI
# =========================================
# This application:
# - Loads CSV orders
# - Calculates shipping totals, refund amounts per recipient
# - Highlights missing values or issues
# - Displays totals and formatting for easy review
# =========================================

Add-Type -AssemblyName System.Windows.Forms        # Load Windows Forms library for GUI
Add-Type -AssemblyName System.Drawing             # Load Drawing library for colors and fonts

# -------------------------------
# Define CSV column headers as variables
# Using variables allows future CSV renaming without breaking code
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
function Update-Refunds($dt) {
    Add-ColumnsIfMissing $dt $colShippingPaid,$colShippingCost,$colTotalShippingPaid,$colRefundAmount  # Ensure all needed columns exist
    Remove-EmptyRows $dt   # Remove empty rows

    # Move Total Shipping Paid column immediately after Shipping Paid
    $dt.Columns[$colTotalShippingPaid].SetOrdinal($dt.Columns[$colShippingPaid].Ordinal + 1)

    # Group rows by recipient
    foreach ($group in ($dt | Group-Object $colRecipient)) {
        $totalPaid = 0; $totalCost = 0

        # Sum Shipping Paid and Shipping Cost for this recipient
        foreach ($r in $group.Group) {
            $paid=0; [decimal]::TryParse($r.$colShippingPaid,[ref]$paid)|Out-Null; $totalPaid += $paid
            $cost=0; [decimal]::TryParse($r.$colShippingCost,[ref]$cost)|Out-Null; $totalCost += $cost
        }

        $first=$true
        foreach ($r in $group.Group) {
            if ($first) {
                # Only first row per recipient gets totals
                $r.$colTotalShippingPaid = if ($totalPaid -ne 0){$totalPaid}else{""}
                $r.$colRefundAmount = if ($totalPaid -ne 0 -or $totalCost -ne 0){$totalPaid - $totalCost}else{""}
                $first=$false
            } else {
                # Other rows empty
                $r.$colTotalShippingPaid = ""; $r.$colRefundAmount = ""
            }
        }
    }
}

# -------------------------------
# GUI Form Setup
# -------------------------------
$form = New-Object Windows.Forms.Form
$form.Text = "NWE Shipping Calculator"          # Window title
$form.Size = '1400,800'                         # Window size
$form.StartPosition = "CenterScreen"           # Centered on screen
$form.Icon = [System.Drawing.SystemIcons]::Application  # Default system icon (can replace with custom later)

# -------------------------------
# Buttons
# -------------------------------
$btnLoad = New-Object Windows.Forms.Button; $btnLoad.Text="Load CSV"; $btnLoad.Size='100,30'
$btnCalc = New-Object Windows.Forms.Button; $btnCalc.Text="Recalc"; $btnCalc.Size='100,30'
$btnSave = New-Object Windows.Forms.Button; $btnSave.Text="Save CSV"; $btnSave.Size='100,30'

# Top panel to hold buttons neatly
$topPanel = New-Object Windows.Forms.FlowLayoutPanel
$topPanel.Height = 50
$topPanel.Top = 10
$topPanel.FlowDirection = 'LeftToRight'
$topPanel.AutoSize = $true
$topPanel.AutoSizeMode = 'GrowAndShrink'
$topPanel.Controls.AddRange(@($btnLoad, $btnCalc, $btnSave))

# Center the topPanel horizontally
$form.Add_Shown({
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
})

# -------------------------------
# DataGridView setup
# -------------------------------
$grid = New-Object Windows.Forms.DataGridView
$grid.AutoGenerateColumns = $true
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font,[System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersHeightSizeMode = 'EnableResizing'
$grid.ColumnHeadersHeight = 30
$grid.RowsDefaultCellStyle.Font = $grid.Font

# Layout: manually position controls
$form.Controls.AddRange(@($topPanel, $grid))

# Position topPanel at top and center
$topPanel.Top = 10
$topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
$form.Add_Shown({
    $topPanel.Left = [math]::Max(0, ($form.ClientSize.Width - $topPanel.Width) / 2)
})

# Position grid below topPanel and size to fill remaining space
$grid.Top = $topPanel.Bottom + 10
$grid.Left = 10
$grid.Width = $form.ClientSize.Width - 20
$grid.Height = $form.ClientSize.Height - $grid.Top - 20
$grid.Anchor = 'Top,Left,Right,Bottom'

# Add grid and panel to form
$form.Controls.AddRange(@($grid,$topPanel))

# Bold font object for later use
$boldFont = New-Object System.Drawing.Font($grid.Font,[System.Drawing.FontStyle]::Bold)

# -------------------------------
# Initialize empty DataTable for grid
# -------------------------------
$emptyTable = New-Object System.Data.DataTable
@($colOrder,$colItemName,$colRecipient,$colQuantity,$colOrderTotal,$colShippingPaid,$colTotalShippingPaid,$colShippingCost,$colRefundAmount) | ForEach-Object { [void]$emptyTable.Columns.Add($_) }
$grid.DataSource = $emptyTable

# ================================
# Cell formatting event: highlights & bold
# ================================
$grid.add_CellFormatting({
    param($src,$e)
    $row = $grid.Rows[$e.RowIndex]
    if (-not $row -or $row.IsNewRow) { return }  # Skip invalid/new rows

    $dt = $grid.DataSource
    $colName = $dt.Columns[$e.ColumnIndex].ColumnName
    $currentRecipient = $row.Cells[$dt.Columns[$colRecipient].Ordinal].Value
    $isFirst = ($e.RowIndex -eq 0 -or $currentRecipient -ne $grid.Rows[$e.RowIndex-1].Cells[$dt.Columns[$colRecipient].Ordinal].Value)

    # First row per recipient gets light blue highlight
    if ($isFirst) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::LightSteelBlue }

    # Red highlight for missing Shipping Paid
    if ($colName -eq $colShippingPaid) {
        $val=$row.Cells[$e.ColumnIndex].Value
        if ([string]::IsNullOrWhiteSpace($val) -or -not [decimal]::TryParse($val,[ref]0) -or ([decimal]$val -eq 0)) {
            $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red
        }
    }

    # Red highlight for missing Shipping Cost (only first row per recipient)
    elseif ($colName -eq $colShippingCost -and $isFirst) {
        $val=$row.Cells[$e.ColumnIndex].Value
        if ([string]::IsNullOrWhiteSpace($val) -or -not [decimal]::TryParse($val,[ref]0)) {
            $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red
        }
    }

    # Refund Amount coloring & bold
    elseif ($colName -eq $colRefundAmount -and $isFirst) {
        $row.Cells[$e.ColumnIndex].Style.Font=$boldFont
        $spVal=$row.Cells[$dt.Columns[$colShippingPaid].Ordinal].Value
        $scVal=$row.Cells[$dt.Columns[$colShippingCost].Ordinal].Value
        $refund=0; [decimal]::TryParse($row.Cells[$e.ColumnIndex].Value,[ref]$refund)|Out-Null
        $hasRed=$false
        if ([string]::IsNullOrWhiteSpace($spVal) -or -not [decimal]::TryParse($spVal,[ref]0) -or ([decimal]$spVal -eq 0)) { $hasRed=$true }
        if ([string]::IsNullOrWhiteSpace($scVal) -or -not [decimal]::TryParse($scVal,[ref]0)) { $hasRed=$true }

        if ($refund -le 0) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Yellow }
        elseif ($hasRed) { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Red }
        else { $row.Cells[$e.ColumnIndex].Style.BackColor=[System.Drawing.Color]::Green }
    }
})

# ================================
# Button Actions
# ================================
# Load CSV
$btnLoad.Add_Click({
    $ofd = New-Object Windows.Forms.OpenFileDialog
    $ofd.Filter="CSV Files (*.csv)|*.csv"
    if ($ofd.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $csv = Import-Csv $ofd.FileName
        $dt = ConvertTo-DataTable $csv
        Remove-EmptyRows $dt
        Add-ColumnsIfMissing $dt $colShippingPaid,$colShippingCost,$colTotalShippingPaid,$colRefundAmount
        $dv = $dt.DefaultView; $dv.Sort="$colRecipient ASC"; $dt=$dv.ToTable()
        Update-Refunds $dt

        # Add totals row at bottom
        $totals = $dt.NewRow()
        if ($dt.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" }
        $totals.$colShippingPaid=($dt|Measure-Object $colShippingPaid -Sum).Sum
        $totals.$colTotalShippingPaid=($dt|Measure-Object $colTotalShippingPaid -Sum).Sum
        $totals.$colShippingCost=($dt|Measure-Object $colShippingCost -Sum).Sum
        $totals.$colRefundAmount=($dt|Measure-Object $colRefundAmount -Sum).Sum
        $dt.Rows.Add($totals)

        $grid.DataSource=$dt; $grid.Refresh()
        $totRow=$grid.Rows.Count-1
        # Make totals row bold
        $grid.Rows[$totRow].Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font=$boldFont } }
    }
})

# Recalculate button
$btnCalc.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        $dt=$grid.DataSource

        # Remove old totals row
        foreach ($r in @($dt.Rows | Where-Object { $_.$colOrder -eq "TOTAL" })) { $dt.Rows.Remove($r) }

        # --- Restore correct order: sort by Recipient, then Order # ---
        $dv = $dt.DefaultView
        $dv.Sort = "$colRecipient ASC, $colOrder ASC"
        $dt = $dv.ToTable()
        $grid.DataSource = $dt

        Update-Refunds $dt

        # Add updated totals row
        $totals = $dt.NewRow()
        if ($dt.Columns.Contains($colOrder)) { $totals.$colOrder="TOTAL" }
        $totals.$colShippingPaid=($dt|Measure-Object $colShippingPaid -Sum).Sum
        $totals.$colTotalShippingPaid=($dt|Measure-Object $colTotalShippingPaid -Sum).Sum
        $totals.$colShippingCost=($dt|Measure-Object $colShippingCost -Sum).Sum
        $totals.$colRefundAmount=($dt|Measure-Object $colRefundAmount -Sum).Sum
        $dt.Rows.Add($totals)
        $grid.Refresh()
        $totRow=$grid.Rows.Count-1
        $grid.Rows[$totRow].Cells | ForEach-Object { if ($_ -and $_.Style) { $_.Style.Font=$boldFont } }
    }
})

# Save CSV button
$btnSave.Add_Click({
    if ($grid.DataSource -is [System.Data.DataTable]) {
        $sfd=New-Object Windows.Forms.SaveFileDialog; $sfd.Filter="CSV Files (*.csv)|*.csv"; $sfd.FileName="ShippingRefunds.csv"
        if ($sfd.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            $grid.DataSource | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV saved successfully.","Saved",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No data to save.","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ================================
# Run the form
# ================================
$form.ShowDialog()
