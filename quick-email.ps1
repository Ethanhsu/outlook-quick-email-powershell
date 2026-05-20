# Quick Email - Outlook COM Automation with WinForms GUI
# Run with: powershell -ExecutionPolicy Bypass -File quick-email.ps1
#
# LAYOUT MATH (absolute positioning, NO TableLayoutPanel/FlowLayoutPanel):
#   Form width  = LM + ctrlW + RM = 16 + 400 + 16 = 432px
#   Form height = 200px (ClientSize, excludes title bar)
#
#   Y positions:
#     y=14   SPL label      (h=18,  fits "SPL Entry")
#     y=36   SPL textbox    (h=24)
#     y=64   Date label     (h=18,  fits "Date")
#     y=86   Date dropdown  (h=24)
#     y=114  Preview label  (h=18,  fits "Preview:")
#     y=136  Preview text   (h=36,  wraps long subject lines)
#     y=180  button         (h=36)
#     y=216  bottom of form
#
#   Gap between rows = 36-14-18 = 4px for label-to-input, etc.
#   Total content height = 136+36+12 = 184, fits in 200px form with room to spare.
#   Using Anchor=Left|Right so controls stretch if form is resized.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Config ----------
$fixedRecipient = "recipient@yourcompany.com"  # <-- CHANGE THIS

function Get-DateOptions {
    $today = Get-Date
    $candidates = New-Object System.Collections.ArrayList
    for ($i = -6; $i -le 1; $i++) {
        # Use .AddMonths() consistently — avoids bug when month=12 and you AddMonths(1)
        $d = $today.AddMonths($i)
        $first = $d.AddDays(1 - $d.Day)        # 1st of month
        $fifth  = $d.AddDays(5 - $d.Day)        # 5th of month
        [void]$candidates.Add($first.ToString("yyyy-MM-dd"))
        [void]$candidates.Add($fifth.ToString("yyyy-MM-dd"))
    }
    $candidates.Sort()
    [array]::Reverse($candidates)
    return $candidates
}

# ---------- Build GUI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Quick Email"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.AutoSize = $false
$form.Dock = "None"

# Layout constants
$LM    = 16          # left margin
$ctrlW = 400         # all controls same width
$formW = 432         # LM + ctrlW + RM
$formH = 216         # total form height (ClientSize, excludes title bar)
$form.ClientSize = New-Object System.Drawing.Size($formW, $formH)

# Y positions (absolute)
$ySPLLabel    = 14
$ySPLBox      = 36
$yDateLabel   = 64
$yDateCombo   = 86
$yPreviewLbl  = 114
$yPreviewTxt  = 136
$yButton      = 180

# Control heights
$hLabel   = 18
$hInput   = 24
$hPreview = 36
$hButton  = 36

# Fonts
$font       = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold8  = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)

# Watermark
$watermark       = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$watermarkColor  = [System.Drawing.Color]::FromArgb(130, 130, 130)
$normalColor     = [System.Drawing.Color]::Black

# ============================================================
# SPL Label
# ============================================================
$splLabel = New-Object System.Windows.Forms.Label
$splLabel.Location = New-Object System.Drawing.Point($LM, $ySPLLabel)
$splLabel.Size     = New-Object System.Drawing.Size($ctrlW, $hLabel)
$splLabel.Text    = "SPL Entry"
$splLabel.Font    = $font
$splLabel.AutoSize = $false
$splLabel.Dock    = "None"
$form.Controls.Add($splLabel)

# ============================================================
# SPL TextBox
# ============================================================
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Location = New-Object System.Drawing.Point($LM, $ySPLBox)
$splBox.Size     = New-Object System.Drawing.Size($ctrlW, $hInput)
$splBox.Font    = $font
$splBox.Text    = $watermark
$splBox.ForeColor = $watermarkColor
$splBox.Anchor  = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($splBox)

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

# ============================================================
# Date Label
# ============================================================
$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Location = New-Object System.Drawing.Point($LM, $yDateLabel)
$dateLabel.Size     = New-Object System.Drawing.Size($ctrlW, $hLabel)
$dateLabel.Text    = "Date"
$dateLabel.Font    = $font
$dateLabel.AutoSize = $false
$dateLabel.Dock    = "None"
$form.Controls.Add($dateLabel)

# ============================================================
# Date ComboBox
# ============================================================
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point($LM, $yDateCombo)
$dateCombo.Size     = New-Object System.Drawing.Size($ctrlW, $hInput)
$dateCombo.Font    = $font
$dateCombo.DropDownStyle = "DropDownList"
$dateCombo.Anchor  = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
foreach ($opt in Get-DateOptions) { [void]$dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

# ============================================================
# Preview Label
# ============================================================
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Location = New-Object System.Drawing.Point($LM, $yPreviewLbl)
$previewLbl.Size     = New-Object System.Drawing.Size($ctrlW, $hLabel)
$previewLbl.Text    = "Preview:"
$previewLbl.Font    = $fontBold8
$previewLbl.AutoSize = $false
$previewLbl.Dock    = "None"
$form.Controls.Add($previewLbl)

# ============================================================
# Preview Text (shows live subject)
# ============================================================
$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name  = "subjectPreview"
$subjectPreview.Location = New-Object System.Drawing.Point($LM, $yPreviewTxt)
$subjectPreview.Size     = New-Object System.Drawing.Size($ctrlW, $hPreview)
$subjectPreview.Text    = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Font    = $font
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.AutoSize = $false
$subjectPreview.Dock    = "None"
$subjectPreview.Anchor  = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($subjectPreview)

# ============================================================
# Create Email Button
# ============================================================
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Location = New-Object System.Drawing.Point($LM, $yButton)
$createBtn.Size     = New-Object System.Drawing.Size($ctrlW, $hButton)
$createBtn.Text     = "Create Email"
$createBtn.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.FlatStyle = "Flat"
$createBtn.Anchor   = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($createBtn)

# ============================================================
# Update Preview on text/selection change
# ============================================================
function Update-Preview {
    $spl = $splBox.Text.Trim()
    if ($spl -eq "" -or $spl -eq $watermark) { $spl = "<SPL Entry>" }
    $dateVal = $dateCombo.SelectedItem
    $subjectPreview.Text = "[Power Automate Admin] Add SPL entry $spl<::>$dateVal"
}
$splBox.Add_TextChanged({ Update-Preview })
$dateCombo.Add_SelectedIndexChanged({ Update-Preview })

# ============================================================
# Button Action — create Outlook mail and display it
# ============================================================
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

# ============================================================
# Icon
# ============================================================
$iconPath = Join-Path $PSScriptRoot "icon.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
}

$form.ShowDialog()