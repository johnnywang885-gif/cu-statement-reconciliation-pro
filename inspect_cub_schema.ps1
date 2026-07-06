param([string]$CubPath = "CUB.MDB", [string]$CubPassword = "thifincub")

try { $dbe = New-Object -ComObject DAO.DBEngine.120 } catch { Write-Error "需安裝 Access Database Engine 2016 (32-bit)"; exit 1 }

$conn = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }
$db = $dbe.OpenDatabase((Resolve-Path $CubPath).Path, $false, $true, $conn)

function Show-Table($name) {
    if ($db.TableDefs | Where-Object { $_.Name -eq $name }) {
        Write-Host "`n=== $name 欄位 ===" -ForegroundColor Cyan
        $td = $db.TableDefs[$name]
        foreach ($f in $td.Fields) {
            $tn = switch($f.Type){
                1{'Boolean'};2{'Byte'};3{'Integer'};4{'Long'};5{'Currency'}
                6{'Single'};7{'Double'};8{'DateTime'};9{'Binary'};10{'Text'}
                11{'LongBinary'};12{'Memo'};13{'GUID'};default{'Unknown'}
            }
            Write-Host ("  {0,-25} {1,-12} Size:{2}" -f $f.Name, $tn, $f.Size)
        }
        # 樣本資料
        $rs = $db.OpenRecordset("SELECT TOP 3 * FROM [$name]")
        if (-not $rs.EOF) {
            Write-Host "  樣本:" -ForegroundColor Gray
            for($i=0;$i -lt 3 -and -not $rs.EOF;$i++) {
                $vals = @()
                foreach($f in $rs.Fields) { $vals += "$($f.Name)=$($f.Value)" }
                Write-Host "    $($vals -join ', ')" -ForegroundColor Gray
                $rs.MoveNext()
            }
        }
        $rs.Close()
    } else {
        Write-Host "`n=== $name：不存在 ===" -ForegroundColor Red
    }
}

# 核心表
Show-Table "LGR"
Show-Table "BOROW"
Show-Table "SER"

# 可能的輔助表
foreach($t in @("RATE","GRP","STAFF","EMP","CLERK","TELLER","WOFF","RECOVER","k_para","PARA")) {
    Show-Table $t
}

$db.Close()
Write-Host "`n完成" -ForegroundColor Green
