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
$form.Size = New-Object System.Drawing.Size(400, 240)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$labelStyle = New-Object System.Windows.Forms.Label
$labelStyle.Text = "Subject Prefix"
$labelStyle.Location = New-Object System.Drawing.Point(20, 20)
$labelStyle.AutoSize = $true
$form.Controls.Add($labelStyle)

$subjectCombo = New-Object System.Windows.Forms.ComboBox
$subjectCombo.Location = New-Object System.Drawing.Point(20, 42)
$subjectCombo.Size = New-Object System.Drawing.Size(360, 25)
$subjectCombo.Items.AddRange(@("[BUG]", "[FEATURE]", "[URGENT]"))
$subjectCombo.SelectedIndex = 0
$form.Controls.Add($subjectCombo)

$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Text = "Date"
$labelDate.Location = New-Object System.Drawing.Point(20, 75)
$labelDate.AutoSize = $true
$form.Controls.Add($labelDate)

$dateCombo = New-Object System.Windows.Forms.ComboBox
$dateCombo.Location = New-Object System.Drawing.Point(20, 97)
$dateCombo.Size = New-Object System.Drawing.Size(360, 25)
foreach ($opt in Get-DateOptions) { $dateCombo.Items.Add($opt) }
$dateCombo.SelectedIndex = 0
$form.Controls.Add($dateCombo)

$labelContent = New-Object System.Windows.Forms.Label
$labelContent.Text = "Content"
$labelContent.Location = New-Object System.Windows.Forms.Point(20, 130)
$labelContent.AutoSize = $true
$form.Controls.Add($labelContent)

$contentBox = New-Object System.Windows.Forms.TextBox
$contentBox.Location = New-Object System.Drawing.Point(20, 152)
$contentBox.Size = New-Object System.Drawing.Size(360, 22)
$form.Controls.Add($contentBox)

$createBtn = New-Object System.Windows.Forms.Button
$createBtn.Text = "Create Email"
$createBtn.Location = New-Object System.Drawing.Point(20, 185)
$createBtn.Size = New-Object System.Drawing.Size(360, 32)
$createBtn.FlatStyle = "Flat"
$createBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createBtn.ForeColor = [System.Drawing.Color]::White
$createBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($createBtn)

# ---------- Button Action ----------
$createBtn.Add_Click({
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)  # 0 = olMailItem

        $subjectType = $subjectCombo.SelectedItem
        $dateVal = $dateCombo.SelectedItem
        $contentVal = $contentBox.Text.Trim()

        $mail.To = $fixedRecipient
        $mail.Subject = "$subjectType $dateVal - $contentVal"
        $mail.Body = "Type: $($subjectType.Replace('[','').Replace(']',''))`nDate: $dateVal`nNote: $contentVal"
        $mail.HTMLBody = "<b>Type:</b> $($subjectType.Replace('[','').Replace(']',''))<br><b>Date:</b> $dateVal<br><b>Note:</b> $contentVal"
        $mail.Display()

        $form.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Quick Email Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.ShowDialog()