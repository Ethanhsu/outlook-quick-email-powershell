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
$form.Size = New-Object System.Drawing.Size(600, 340)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$labelX = 16
$inputX = 16
$inputW = 568
$rowH = 30
$rowGap = 6
$y = 20

# SPL Entry (user text input)
$lblSpl = New-Object System.Windows.Forms.Label
$lblSpl.Text = "SPL Entry"
$lblSpl.Location = New-Object System.Drawing.Point($labelX, $y)
$lblSpl.AutoSize = $true
$form.Controls.Add($lblSpl)

$y += 18
$splBox = New-Object System.Windows.Forms.TextBox
$splBox.Location = New-Object System.Drawing.Point($inputX, $y)
$splBox.Size = New-Object System.Drawing.Size($inputW, 26)
$splBox.PlaceholderText = "e.g. 14-41-13.00-UG-U00-STD-HEL-04/84"
$splBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($splBox)

$y += $rowH + $rowGap

# Date
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Date"
$lblDate.Location = New-Object System.Drawing.Point($labelX, $y)
$lblDate.AutoSize = $true
$form.Controls.Add($lblDate)

$y += 18
$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point($inputX, $y)
$dateCombo.Size = New-Object System.Drawing.Size($inputW, 26)
$dateCombo.Font = New-Object System.Drawing.Font("Segoe UI", 10)
foreach ($opt in Get-DateOptions) { $dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

$y += $rowH + 12

# Preview label
$previewLbl = New-Object System.Windows.Forms.Label
$previewLbl.Text = "Preview:"
$previewLbl.Location = New-Object System.Drawing.Point($labelX, $y)
$previewLbl.AutoSize = $true
$previewLbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($previewLbl)

$y += 18

# Preview subject
$subjectPreview = New-Object System.Windows.Forms.Label
$subjectPreview.Name = "subjectPreview"
$subjectPreview.Text = "[Power Automate Admin] Add SPL entry <::>"
$subjectPreview.Location = New-Object System.Drawing.Point($labelX, $y)
$subjectPreview.AutoSize = $true
$subjectPreview.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$subjectPreview.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($subjectPreview)

$y += $rowH + 16

# Create button
$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Location = New-Object System.Drawing.Point($inputX, $y)
$createBtn.Size = New-Object System.Drawing.Size($inputW, 40)
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
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
        $mail = $outlook.CreateItem(0)  # 0 = olMailItem

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

# Icon - set to Outlook's icon so it looks professional
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")

$form.ShowDialog()