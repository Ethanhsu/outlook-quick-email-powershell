# Quick Email - Outlook COM Automation with WinForms GUI
# Run with: powershell -ExecutionPolicy Bypass -File quick-email.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Config ----------
$fixedRecipient = "recipient@yourcompany.com"  # <-- CHANGE THIS

function Get-DateOptions {
    $today = Get-Date
    $candidates = New-Object System.Collections.ArrayList
    for ($i = -6; $i -le 1; $i++) {
        $d = $today.AddMonths($i)
        $first = Get-Date -Year $d.Year -Month $d.Month -Day 1
        $fifth = Get-Date -Year $d.Year -Month $d.Month -Day 5
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

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$watermark = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$LM = 16
$RM = 16
$ctrlW = 400
$formW = $LM + $ctrlW + $RM  # 432

# Form height: row heights + top/bottom padding
# Row heights: 16+24+10+16+24+10+16+40+10+36 = 202
# Top+bottom padding: 14+14 = 28
$formH = 230
$form.ClientSize = New-Object System.Drawing.Size($formW, $formH)

# Single column TableLayoutPanel — stretches controls horizontally
$tbl = New-Object System.Windows.Forms.TableLayoutPanel
$tbl.Dock = "Fill"
$tbl.AutoSize = $false
$tbl.ColumnCount = 1
$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tbl.RowCount = 0
$tbl.Padding = New-Object System.Windows.Forms.Padding($LM, 14, $RM, 14)

function AddRow($h) {
    $tbl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $h)))
    $tbl.RowCount++
    return $tbl.RowCount - 1
}

function Place($ctrl, $row) {
    $tbl.SetColumn($ctrl, 0)
    $tbl.SetRow($ctrl, $row)
    $tbl.Controls.Add($ctrl)
}

$r0 = AddRow 16     # SPL label
$r1 = AddRow 24     # SPL input
$r2 = AddRow 10     # gap
$r3 = AddRow 16     # Date label
$r4 = AddRow 24     # Date dropdown
$r5 = AddRow 10     # gap
$r6 = AddRow 16     # Preview label
$r7 = AddRow 40     # Preview text
$r8 = AddRow 10     # gap
$r9 = AddRow 36     # button

# SPL label
$splLabel = New-Object System.Windows.Forms.Label
$splLabel.Text = "SPL Entry"
$splLabel.Font = $font
$splLabel.Dock = "Fill"
Place $splLabel $r0

# SPL input — stretches to full TableLayoutPanel width
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Font = $font
$splBox.Size = New-Object System.Drawing.Size($ctrlW, 22)
$splBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$splBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Left

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
Place $splBox $r1

# Date label
$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Text = "Date"
$dateLabel.Font = $font
$dateLabel.Dock = "Fill"
Place $dateLabel $r3

# Date dropdown
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Font = $font
$dateCombo.Size = New-Object System.Drawing.Size($ctrlW, 22)
$dateCombo.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
foreach ($opt in Get-DateOptions) { [void]$dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
Place $dateCombo $r4

# Preview label
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$previewLbl.Dock = "Fill"
Place $previewLbl $r6

# Preview text
$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Size = New-Object System.Drawing.Size($ctrlW, 40)
$subjectPreview.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = $font
$subjectPreview.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Place $subjectPreview $r7

# Button
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Size = New-Object System.Drawing.Size($ctrlW, 36)
$createBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
Place $createBtn $r9

$form.Controls.Add($tbl)

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