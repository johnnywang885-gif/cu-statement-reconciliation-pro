# get_password.ps1 — 彈出圖形化視窗輸入 CUB.MDB 密碼
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CUB.MDB 密碼'
$form.Size = New-Object System.Drawing.Size(350, 150)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = '請輸入 CUB.MDB 密碼：'
$lbl.Location = New-Object System.Drawing.Point(20, 20)
$lbl.AutoSize = $true

$txt = New-Object System.Windows.Forms.TextBox
$txt.PasswordChar = '*'
$txt.Location = New-Object System.Drawing.Point(20, 45)
$txt.Size = New-Object System.Drawing.Size(290, 25)

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = '確定'
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$btnOk.Location = New-Object System.Drawing.Point(140, 80)
$btnOk.Size = New-Object System.Drawing.Size(80, 25)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = '取消'
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$btnCancel.Location = New-Object System.Drawing.Point(230, 80)
$btnCancel.Size = New-Object System.Drawing.Size(80, 25)

$form.Controls.Add($lbl)
$form.Controls.Add($txt)
$form.Controls.Add($btnOk)
$form.Controls.Add($btnCancel)
$form.AcceptButton = $btnOk
$form.CancelButton = $btnCancel

# Bring form to front
$form.Add_Shown({$form.Activate()})

if ($form.ShowDialog() -eq 'OK') {
    Write-Output $txt.Text
} else {
    Write-Output ''
}
