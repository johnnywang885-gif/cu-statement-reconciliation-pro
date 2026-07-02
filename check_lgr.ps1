param([string]$CubPath = "CUB.MDB", [string]$CubPassword = "thifincub")
$dbe = New-Object -ComObject DAO.DBEngine.120
$db = $dbe.OpenDatabase((Resolve-Path $CubPath).Path, $false, $true, ";PWD=$CubPassword")

Write-Host "=== ALL TABLES ==="
foreach ($t in $db.TableDefs) {
    $n = $t.Name
    if (-not $n.StartsWith("MSys") -and -not $n.StartsWith("~")) {
        Write-Host $n
    }
}

$found = $false
foreach ($t in $db.TableDefs) {
    if ($t.Name -eq "LGR") { $found = $true }
}
if ($found) {
    Write-Host "`n=== LGR FIELDS ==="
    $td = $db.TableDefs["LGR"]
    foreach ($f in $td.Fields) {
        Write-Host ("{0,-25} Type={1,-5} Size={2}" -f $f.Name, $f.Type, $f.Size)
    }
    $rs = $db.OpenRecordset("SELECT TOP 2 * FROM LGR")
    if (-not $rs.EOF) {
        Write-Host "SAMPLES:"
        for ($i = 0; $i -lt 2 -and -not $rs.EOF; $i++) {
            $line = ""
            foreach ($f in $rs.Fields) {
                if ($line.Length -gt 0) { $line += "|" }
                $line += ("{0}={1}" -f $f.Name, $f.Value)
            }
            Write-Host $line
            $rs.MoveNext()
        }
    }
    $rs.Close()
} else {
    Write-Host "LGR NOT FOUND"
}

Write-Host "`n=== BOROW USR SAMPLES ==="
$rs = $db.OpenRecordset("SELECT DISTINCT TOP 8 USR FROM BOROW WHERE USR IS NOT NULL AND USR != ''")
while (-not $rs.EOF) {
    Write-Host ("USR=[{0}]" -f $rs.Fields("USR").Value)
    $rs.MoveNext()
}
$rs.Close()

Write-Host "`n=== SER OUTDAT COUNT ==="
$rs = $db.OpenRecordset("SELECT Count(*) AS Cnt, OUTDAT FROM SER GROUP BY OUTDAT ORDER BY Count(*) DESC")
while (-not $rs.EOF) {
    $v = [string]$rs.Fields("OUTDAT").Value
    $c = $rs.Fields("Cnt").Value
    if ($v.Length -eq 0) { $v = "(empty)" }
    Write-Host ("OUTDAT=[{0}] count={1}" -f $v, $c)
    $rs.MoveNext()
}
$rs.Close()

$db.Close()
