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
        $d = (Get-Date -Year $today.Year -Month ($today.Month + $i) -Day 1)
        $candidates += $d.ToString("yyyy-MM-dd")
        $candidates += (Get-Date -Year $d.Year -Month $d.Month -Day 5).ToString("yyyy-MM-dd")
    }
    $candidates = $candidates | Sort-Object -Descending
    return $candidates
}

# ---------- Build GUI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Quick Email"
$form.Size = New-Object System.Drawing.Size(460, 295)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelX = 16
$inputW = 428
$inputH = 26
$rowH = 26
$gapLabel = 5
$gapInput = 10
$y = 16

# SPL Entry
$lblSpl = New-Object System.Windows.Forms.Label
$lblSpl.Text = "SPL Entry"
$lblSpl.Location = New-Object System.Drawing.Point($labelX, $y)
$lblSpl.Font = $font
$lblSpl.AutoSize = $true
$form.Controls.Add($lblSpl)

$y += 14
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Location = New-Object System.Drawing.Point($labelX, $y)
$splBox.Size = New-Object System.Drawing.Size($inputW, $inputH)
$splBox.PlaceholderText = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$splBox.Font = $font
$form.Controls.Add($splBox)

$y += $rowH + $gapInput

# Date
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Date"
$lblDate.Location = New-Object System.Drawing.Point($labelX, $y)
$lblDate.Font = $font
$lblDate.AutoSize = $true
$form.Controls.Add($lblDate)

$y += 14
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point($labelX, $y)
$dateCombo.Size = New-Object System.Drawing.Size($inputW, $inputH)
$dateCombo.Font = $font
foreach ($opt in Get-DateOptions) { $dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

$y += $rowH + $gapInput + 2

# Preview
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Location = New-Object System.Drawing.Point($labelX, $y)
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$previewLbl.AutoSize = $true
$form.Controls.Add($previewLbl)

$y += 14

$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Location = New-Object System.Drawing.Point($labelX, $y)
$subjectPreview.MaximumSize = New-Object System.Drawing.Size($inputW, 0)  # wrap at input width
$subjectPreview.AutoSize = $true
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = $font
$form.Controls.Add($subjectPreview)

$y += $rowH + 10

# Create button
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Location = New-Object System.Drawing.Point($labelX, $y)
$createBtn.Size = New-Object System.Drawing.Size($inputW, 36)
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($createBtn)

# ---------- Update preview on change ----------
function Update-Preview {
    $spl = $splBox.Text.Trim()
    if ($spl -eq "") { $spl = "<SPL Entry>" }
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