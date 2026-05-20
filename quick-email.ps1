# Quick Email - Outlook COM Automation with WinForms GUI
# Run with: powershell -ExecutionPolicy Bypass -File quick-email.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Config ----------
$fixedRecipient = "recipient@yourcompany.com"  # <-- CHANGE THIS

# Generate 1st and 5th of each month, 6 months ago to 1 month ahead, descending
function Get-DateOptions {
    $today = Get-Date
    $candidates = @()
    for ($i = -6; $i -le 1; $i++) {
        $d = $today.AddMonths($i)
        $first = Get-Date -Year $d.Year -Month $d.Month -Day 1
        $fifth = Get-Date -Year $d.Year -Month $d.Month -Day 5
        $candidates += $first.ToString("yyyy-MM-dd")
        $candidates += $fifth.ToString("yyyy-MM-dd")
    }
    $candidates = $candidates | Sort-Object -Descending
    return $candidates
}

# ---------- Build GUI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Quick Email"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$watermark = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"

# FlowLayoutPanel with WrapContents=false goes straight down
$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.AutoSize = $true
$flow.AutoSizeMode = "GrowOnly"
$flow.FlowDirection = "TopDown"
$flow.WrapContents = $false
$flow.Padding = New-Object System.Windows.Forms.Padding(16, 14, 16, 14)
$flow.Dock = "Fill"

# Each row is a Panel with FlowDirection=LeftToRight, so label stays left, control fills right
function New-RowPanel {
    $p = New-Object System.Windows.Forms.Panel
    $p.AutoSize = $true
    $p.AutoSizeMode = "GrowWidth"
    $p.FlowDirection = "TopDown"
    $p.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
    return $p
}

# ---------- SPL Entry row ----------
$rowSpl = New-RowPanel

$splLabel = New-Object System.Windows.Forms.Label
$splLabel.Text = "SPL Entry"
$splLabel.Font = $font
$splLabel.AutoSize = $true
$splLabel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 3)

$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Width = 398
$splBox.Font = $font
$splBox.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

$watermarkColor = [System.Drawing.Color]::FromArgb(130, 130, 130)
$normalColor = [System.Drawing.Color]::Black
$splBox.ForeColor = $watermarkColor
$splBox.Text = $watermark

$splBox.Add_GotFocus({
    if ($this.Text -eq $global:watermark) {
        $this.Text = ""
        $this.ForeColor = $global:normalColor
    }
})
$splBox.Add_LostFocus({
    if ($this.Text -eq "") {
        $this.Text = $global:watermark
        $this.ForeColor = $global:watermarkColor
    }
})

$rowSpl.Controls.Add($splLabel)
$rowSpl.Controls.Add($splBox)

# ---------- Date row ----------
$rowDate = New-RowPanel

$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Text = "Date"
$dateLabel.Font = $font
$dateLabel.AutoSize = $true
$dateLabel.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 3)

$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Width = 398
$dateCombo.Font = $font
foreach ($opt in Get-DateOptions) { [void]$dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0

$rowDate.Controls.Add($dateLabel)
$rowDate.Controls.Add($dateCombo)

# ---------- Preview row ----------
$rowPreview = New-RowPanel
$rowPreview.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)

$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$previewLbl.AutoSize = $true

$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.AutoSize = $false
$subjectPreview.Width = 398
$subjectPreview.Height = 36
$subjectPreview.TextAlign = "TopLeft"
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = $font

$rowPreview.Controls.Add($previewLbl)
$rowPreview.Controls.Add($subjectPreview)

# ---------- Button row ----------
$rowBtn = New-RowPanel
$rowBtn.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)

$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Width = 398
$createBtn.Height = 36
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$rowBtn.Controls.Add($createBtn)

# Add all rows to flow panel
$flow.Controls.Add($rowSpl)
$flow.Controls.Add($rowDate)
$flow.Controls.Add($rowPreview)
$flow.Controls.Add($rowBtn)

$form.Controls.Add($flow)
$form.AutoSize = $true
$form.AutoSizeMode = "GrowAndShrink"
$form.MinimumSize = New-Object System.Drawing.Size(430, 200)

# ---------- Update preview on change ----------
function Update-Preview {
    $spl = $splBox.Text.Trim()
    if ($spl -eq "" -or $spl -eq $watermark) { $spl = "<SPL Entry>" }
    $dateVal = $dateCombo.SelectedItem
    $subjectPreview.Text = "[Power Automate Admin] Add SPL entry $spl<::>$dateVal"
}
$splBox.Add_TextChanged({ Update-Preview })
$dateCombo.Add_SelectedIndexChanged({ Update-Preview })

# ---------- Button Action ----------
$createBtn.Add_Click({
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $spl = $splBox.Text.Trim()
        if ($spl -eq $watermark) { $spl = "" }
        $dateVal = $dateCombo.SelectedItem
        if ($spl -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Please enter the SPL Entry.", "Quick Email", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $fullSubject = "[Power Automate Admin] Add SPL entry $spl<::>$dateVal"
        $mail.To = $fixedRecipient
        $mail.Subject = $fullSubject
        $mail.Display()
        $form.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Quick Email Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Icon
$iconPath = Join-Path $PSScriptRoot "icon.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
}

$form.ShowDialog()