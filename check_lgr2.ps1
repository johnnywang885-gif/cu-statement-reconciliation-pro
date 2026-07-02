$dbe = New-Object -ComObject DAO.DBEngine.120
$db = $dbe.OpenDatabase((Resolve-Path "CUB.MDB").Path, $false, $true, ";PWD=thifincub")

$found = $false
foreach ($t in $db.TableDefs) {
    $n = $t.Name
    if (-not $n.StartsWith("MSys")) {
        Write-Host $n
        if ($n.ToUpper() -eq "LGR") { $found = $true }
    }
}

if ($found) {
    Write-Host "`nLGR EXISTS!"
    $rs = $db.OpenRecordset("SELECT Count(*) AS C FROM LGR")
    Write-Host ("LGR records: {0}" -f $rs.Fields("C").Value)
    $rs.Close()
    
    $rs = $db.OpenRecordset("SELECT TOP 3 * FROM LGR")
    Write-Host "Sample LGR columns:"
    foreach ($f in $rs.Fields) {
        Write-Host ("  {0}" -f $f.Name)
    }
    $rs.Close()
} else {
    Write-Host "`nLGR NOT FOUND in table list"
}

$db.Close()
