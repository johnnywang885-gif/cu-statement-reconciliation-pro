Set-StrictMode -Version Latest

$script:AnomalyWeights = [ordered]@{
    DiffShare            = 5
    DiffLoan             = 5
    DiffReserve          = 5
    NonMember            = 5
    SameAddrMul          = 2
    RelatedParty         = 3
    NewJoinLoan          = 2
    RecDiscrepancy       = 2
    AuditOverdue         = 2
    ExceedPayDate        = 1
    DirectorOverLimit    = 3
    DuplicateLoan        = 2
    NegativeLoanBalance  = 4
    LoanBeforeApproval   = 4
    DormantActivation    = 4
    NewAccountBurst      = 3
    RoundAmountTxn       = 2
}

function Get-MemberValue {
    param(
        [hashtable] $Hash,
        [string] $Key,
        $Default = 0
    )
    if ($Hash.ContainsKey($Key)) {
        $v = $Hash[$Key]
        if ($null -eq $v) { return $Default }
        return $v
    }
    return $Default
}

function Get-AnomalyScore {
    param(
        [Parameter(Mandatory)] [hashtable] $Member
    )

    $lgrShare   = [double](Get-MemberValue $Member 'LGR_Share')
    $serShare   = [double](Get-MemberValue $Member 'SER_Share')
    $lgrLoan    = [double](Get-MemberValue $Member 'LGR_Loan')
    $serLoan    = [double](Get-MemberValue $Member 'SER_Loan')
    $lgrReserve = [double](Get-MemberValue $Member 'LGR_Reserve')
    $serReserve = [double](Get-MemberValue $Member 'SER_Reserve')

    $diffShare   = $lgrShare   - $serShare
    $diffLoan    = $lgrLoan    - $serLoan
    $diffReserve = $lgrReserve - $serReserve

    $score = 0
    $flags = New-Object System.Collections.Generic.List[string]

    if ($diffShare -ne 0) {
        $score += $script:AnomalyWeights.DiffShare
        $flags.Add(('股金差{0}' -f $diffShare)) | Out-Null
    }
    if ($diffLoan -ne 0) {
        $score += $script:AnomalyWeights.DiffLoan
        $flags.Add(('貸款差{0}' -f $diffLoan)) | Out-Null
    }
    if ($diffReserve -ne 0) {
        $score += $script:AnomalyWeights.DiffReserve
        $flags.Add(('備轉差{0}' -f $diffReserve)) | Out-Null
    }

    if ([bool](Get-MemberValue $Member 'IsNonMember' $false)) {
        $score += $script:AnomalyWeights.NonMember
        $flags.Add('非社員貸款') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasSameAddressMultiple' $false)) {
        $score += $script:AnomalyWeights.SameAddrMul
        $flags.Add('同地址多戶貸款(>=3)') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasRelatedPartyLoan' $false)) {
        $score += $script:AnomalyWeights.RelatedParty
        $flags.Add('關係人放款(同GRPNO)') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasNewJoinLoan' $false)) {
        $score += $script:AnomalyWeights.NewJoinLoan
        $flags.Add('新入社立即貸款') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasRecDiscrepancy' $false)) {
        $score += $script:AnomalyWeights.RecDiscrepancy
        $flags.Add('借據與審查紀錄不符') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasAuditOverdue' $false)) {
        $score += $script:AnomalyWeights.AuditOverdue
        $flags.Add('審查後逾期未入帳') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasExceedPayDate' $false)) {
        $score += $script:AnomalyWeights.ExceedPayDate
        $flags.Add('超過約定還款日期') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasDirectorOverLimit' $false)) {
        $score += $script:AnomalyWeights.DirectorOverLimit
        $flags.Add('董監事貸款超限') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasDuplicateLoan' $false)) {
        $score += $script:AnomalyWeights.DuplicateLoan
        $flags.Add('重複交易') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasDormantActivation' $false)) {
        $score += $script:AnomalyWeights.DormantActivation
        $flags.Add('休眠戶激活') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasRoundAmountTxn' $false)) {
        $score += $script:AnomalyWeights.RoundAmountTxn
        $flags.Add('整數金額(>=50%)') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasNewAccountBurst' $false)) {
        $score += $script:AnomalyWeights.NewAccountBurst
        $flags.Add('新帳戶短期爆發') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasNegativeLoanBalance' $false)) {
        $score += $script:AnomalyWeights.NegativeLoanBalance
        $flags.Add('貸款結餘為負') | Out-Null
    }
    if ([bool](Get-MemberValue $Member 'HasLoanBeforeApproval' $false)) {
        $score += $script:AnomalyWeights.LoanBeforeApproval
        $flags.Add('先放後審') | Out-Null
    }

    return [PSCustomObject]@{
        Score       = $score
        Flags       = $flags.ToArray()
        Flags_cn    = ($flags -join '; ')
        DiffShare   = $diffShare
        DiffLoan    = $diffLoan
        DiffReserve = $diffReserve
    }
}

function Get-ScoreCategory {
    param([int]$Score)

    if ($Score -ge 10) { return 'High' }
    if ($Score -ge 5)  { return 'Mid'  }
    if ($Score -gt 0)  { return 'Low'  }
    return 'None'
}

function Test-AnomalyShouldInclude {
    param([int]$Score)
    return $Score -gt 0
}

function Get-TopPercentCount {
    param(
        [Parameter(Mandatory)] [int] $Total,
        [Parameter(Mandatory)] [int] $Percent
    )

    if ($Total -le 0) { return 0 }
    if ($Percent -lt 1 -or $Percent -gt 100) {
        throw "Percent must be between 1 and 100 (got $Percent)"
    }
    return [Math]::Max(1, [int]($Total * $Percent / 100))
}

function Get-ResultCsvName {
    param(
        [string] $Prefix = 'CUB_異常社員_對帳單排序_',
        [datetime] $When = (Get-Date)
    )
    $ts = $When.ToString('yyyyMMdd_HHmmss')
    return "$Prefix$ts.csv"
}

function Get-AnomalyWeights {
    $copy = [ordered]@{}
    foreach ($k in $script:AnomalyWeights.Keys) {
        $copy[$k] = $script:AnomalyWeights[$k]
    }
    return $copy
}

Export-ModuleMember -Function @(
    'Get-AnomalyScore'
    'Get-ScoreCategory'
    'Test-AnomalyShouldInclude'
    'Get-TopPercentCount'
    'Get-ResultCsvName'
    'Get-AnomalyWeights'
)
