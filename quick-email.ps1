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
$form.Size = New-Object System.Drawing.Size(440, 250)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$leftX = 22
$inputW = 396
$inputH = 24

# SPL Entry label
$lblSpl = New-Object System.Windows.Forms.Label
$lblSpl.Text = "SPL Entry"
$lblSpl.Location = New-Object System.Drawing.Point($leftX, 12)
$lblSpl.Size = New-Object System.Drawing.Size(80, 14)
$lblSpl.Font = $font
$form.Controls.Add($lblSpl)

# SPL Entry input
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Location = New-Object System.Drawing.Point($leftX, 26)
$splBox.Size = New-Object System.Drawing.Size($inputW, $inputH)
$splBox.PlaceholderText = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$splBox.Font = $font
$form.Controls.Add($splBox)

# Date label
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Date"
$lblDate.Location = New-Object System.Drawing.Point($leftX, 56)
$lblDate.Size = New-Object System.Drawing.Size(80, 14)
$lblDate.Font = $font
$form.Controls.Add($lblDate)

# Date dropdown
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point($leftX, 70)
$dateCombo.Size = New-Object System.Drawing.Size($inputW, $inputH)
$dateCombo.Font = $font
foreach ($opt in Get-DateOptions) { $dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

# Preview label
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Location = New-Object System.Drawing.Point($leftX, 100)
$previewLbl.Size = New-Object System.Drawing.Size(80, 14)
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($previewLbl)

# Preview subject
$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Location = New-Object System.Drawing.Point($leftX, 114)
$subjectPreview.MaximumSize = New-Object System.Drawing.Size($inputW, 0)
$subjectPreview.AutoSize = $true
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = $font
$form.Controls.Add($subjectPreview)

# Create button
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Location = New-Object System.Drawing.Point($leftX, 148)
$createBtn.Size = New-Object System.Drawing.Size($inputW, 34)
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