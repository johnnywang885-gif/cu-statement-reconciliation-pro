# select_cub.ps1 — 彈出圖形化視窗選擇 CUB.MDB
Add-Type -AssemblyName System.Windows.Forms
$f = New-Object System.Windows.Forms.OpenFileDialog
$f.Title = '請選擇 CUB.MDB'
$f.Filter = 'Access 資料庫 (*.mdb)|*.mdb|所有檔案 (*.*)|*.*'
$f.InitialDirectory = $PSScriptRoot
if ($f.ShowDialog() -eq 'OK') {
    Write-Output $f.FileName
} else {
    Write-Output ''
}
