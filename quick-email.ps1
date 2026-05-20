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
$form.Size = New-Object System.Drawing.Size(430, 265)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$leftX = 16
$formW = 430
$inputW = 398
$inputH = 22
$lblH = 13
$gapLblToInput = 4
$gapInputToNext = 14
$bottomMargin = 16

$y = 16

# --- SPL Entry label ---
$lblSpl = New-Object System.Windows.Forms.Label
$lblSpl.Text = "SPL Entry"
$lblSpl.Location = New-Object System.Drawing.Point($leftX, $y)
$lblSpl.Size = New-Object System.Drawing.Size(80, $lblH)
$lblSpl.Font = $font
$form.Controls.Add($lblSpl)

$y += $lblH + $gapLblToInput

# --- SPL Entry input ---
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Location = New-Object System.Drawing.Point($leftX, $y)
$splBox.Size = New-Object System.Drawing.Size($inputW, $inputH)
$splBox.Text = ""
$splBox.Font = $font
$form.Controls.Add($splBox)

# Watermark via GotFocus/LostFocus
$watermark = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$watermarkColor = [System.Drawing.Color]::FromArgb(130, 130, 130)
$normalColor = $splBox.ForeColor

$splBox.Add_GotFocus({
    if ($this.Text -eq $watermark) {
        $this.Text = ""
        $this.ForeColor = $normalColor
    }
})
$splBox.Add_LostFocus({
    if ($this.Text -eq "") {
        $this.Text = $watermark
        $this.ForeColor = $watermarkColor
    }
})
# Init state
$splBox.Text = $watermark
$splBox.ForeColor = $watermarkColor

$y += $inputH + $gapInputToNext

# --- Date label ---
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Date"
$lblDate.Location = New-Object System.Drawing.Point($leftX, $y)
$lblDate.Size = New-Object System.Drawing.Size(80, $lblH)
$lblDate.Font = $font
$form.Controls.Add($lblDate)

$y += $lblH + $gapLblToInput

# --- Date dropdown ---
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point($leftX, $y)
$dateCombo.Size = New-Object System.Drawing.Size($inputW, $inputH)
$dateCombo.Font = $font
foreach ($opt in Get-DateOptions) { [void]$dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

$y += $inputH + $gapInputToNext

# --- Preview label ---
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Location = New-Object System.Drawing.Point($leftX, $y)
$previewLbl.Size = New-Object System.Drawing.Size(80, $lblH)
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($previewLbl)

$y += $lblH + $gapLblToInput

# --- Preview subject ---
$previewH = 40  # Fixed height for preview text area (allows 2 lines of wrapped text)
$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Location = New-Object System.Drawing.Point($leftX, $y)
$subjectPreview.Size = New-Object System.Drawing.Size($inputW, $previewH)
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = $font
$form.Controls.Add($subjectPreview)

$y += $previewH + $gapInputToNext

# --- Create button ---
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Location = New-Object System.Drawing.Point($leftX, $y)
$createBtn.Size = New-Object System.Drawing.Size($inputW, 34)
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($createBtn)

$y += 34 + 16  # button height + bottom margin

$form.Size = New-Object System.Drawing.Size($formW, $y)

# ---------- Update preview on change ----------
function Update-Preview {
    $spl = $splBox.Text.Trim()
    $isPlaceholder = ($splBox.Text -eq $watermark)
    if ($isPlaceholder) { $spl = "<SPL Entry>" }
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