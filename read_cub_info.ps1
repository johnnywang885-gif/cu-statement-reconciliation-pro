# read_cub_info.ps1 — 32-bit 專用：從 CUB.MDB 讀取合作社名與日期
# 輸出 JSON 格式供主腳本解析
param(
    [string]$CubPath = "",
    [string]$CubPassword = ""
)

if ([string]::IsNullOrEmpty($CubPath)) {
    $CubPath = Join-Path $PSScriptRoot "CUB.MDB"
}
$CubPath = [System.IO.Path]::GetFullPath($CubPath)

if (-not (Test-Path $CubPath)) {
    Write-Output '{"error":"CUB.MDB not found"}'
    exit 1
}

$connectStr = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }

try {
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)

    # 讀取 k_para（日期）
    $rs = $db.OpenRecordset("SELECT Item, Para FROM k_para")
    $kPara = @{}
    while (-not $rs.EOF) {
        $kPara[$rs.Fields["Item"].Value] = $rs.Fields["Para"].Value
        $rs.MoveNext()
    }
    $rs.Close()

    $rawDate = ""
    if ($kPara.ContainsKey("eDate") -and $kPara["eDate"]) { $rawDate = [string]$kPara["eDate"] }

    # 合作社名：從 SER 取 ACCNO='000000' 的 ACCNM
    $coopName = ""
    $rsName = $db.OpenRecordset("SELECT ACCNM FROM SER WHERE ACCNO='000000'")
    if (-not $rsName.EOF) {
        $val = $rsName.Fields["ACCNM"].Value
        if ($val) { $coopName = [string]$val.Trim() }
    }
    $rsName.Close()

    $db.Close()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbe)

    # 輸出 JSON
    $rawDate = $rawDate -replace '"', '\"'
    $coopName = $coopName -replace '"', '\"'
    Write-Output "{`"rawDate`":`"$rawDate`",`"coopName`":`"$coopName`"}"
}
catch {
    Write-Output "{`"error`":`"$($_.Exception.Message -replace '"', '\"')`"}"
    exit 1
}
