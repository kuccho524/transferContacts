# ==========================================================
# GAM 連絡先登録ツール CLI版
# 統合GUIアプリケーションから呼び出される版
# ==========================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$ContactCsvFile,

    [string]$LabelCsvFile = "",

    [Parameter(Mandatory = $true)]
    [string]$TargetUserEmail,

    [string]$TempJsonFile = "$env:TEMP\contact_temp.json",

    [string]$LogDir = "./logs"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function Log-Message {
    param([string]$Message)
    $TimeStamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $LogLine = "[$TimeStamp] $Message"
    Write-Output $LogLine
}

function Invoke-GamCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    try {
        $output = & gam @Arguments 2>&1 | Out-String
        return @{
            ExitCode = $LASTEXITCODE
            Output   = $output.Trim()
        }
    }
    catch {
        throw "GAM実行エラー: $_"
    }
}

function Get-DefaultMethodType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ColumnName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Email", "Phone")]
        [string]$Kind
    )

    if ($Kind -eq "Email") {
        if ($ColumnName -like 'Primary*') { return "work" }
        if ($ColumnName -like 'Secondary*') { return "home" }
        return "other"
    }

    if ($ColumnName -like 'Primary*') { return "mobile" }
    if ($ColumnName -like 'Secondary*') { return "home" }
    return "work"
}

function Get-EmailObjects {
    param([pscustomobject]$Row)

    $emailObjects = @()
    $emailColumns = $Row.PSObject.Properties.Name | Where-Object { $_ -like '*EmailAddress' }

    foreach ($valueColumn in $emailColumns) {
        $value = $Row.$valueColumn
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $labelColumn = $valueColumn -replace 'Address$', 'Label'
        $labelValue = $Row.$labelColumn
        $emailObject = @{
            value = $value
            type  = if ($labelValue) { $labelValue } else { Get-DefaultMethodType -ColumnName $valueColumn -Kind Email }
        }

        if ($valueColumn -like 'Primary*') {
            $emailObject.metadata = @{ primary = $true }
        }

        $emailObjects += $emailObject
    }

    return $emailObjects
}

function Get-PhoneObjects {
    param([pscustomobject]$Row)

    $phoneObjects = @()
    $phoneColumns = $Row.PSObject.Properties.Name | Where-Object { $_ -like '*PhoneNumber' }

    foreach ($valueColumn in $phoneColumns) {
        $value = $Row.$valueColumn
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $labelColumn = $valueColumn -replace 'Number$', 'Label'
        $labelValue = $Row.$labelColumn
        $phoneObject = @{
            value = $value
            type  = if ($labelValue) { $labelValue } else { Get-DefaultMethodType -ColumnName $valueColumn -Kind Phone }
        }

        if ($valueColumn -like 'Primary*') {
            $phoneObject.metadata = @{ primary = $true }
        }

        $phoneObjects += $phoneObject
    }

    return $phoneObjects
}

function Get-LabelColumns {
    param([string]$CsvPath)

    $sampleRow = Import-Csv -Path $CsvPath -Encoding UTF8 | Select-Object -First 1
    if ($sampleRow) {
        return $sampleRow.PSObject.Properties.Name | Where-Object {
            $_ -like '*Label' -and
            $_ -ne 'PrimaryLabel' -and
            $_ -ne 'TargetEmail' -and
            $_ -ne 'ContactID'
        }
    }

    $firstLine = Get-Content -Path $CsvPath -TotalCount 1 -Encoding UTF8
    if (-not $firstLine) {
        return @()
    }

    return ($firstLine -split ',') | Where-Object {
        $_ -like '*Label' -and
        $_ -ne 'PrimaryLabel' -and
        $_ -ne 'TargetEmail' -and
        $_ -ne 'ContactID'
    }
}

try {
    Log-Message "=== 連絡先登録処理開始 ==="
    Log-Message "Target User: $TargetUserEmail"
    Log-Message "Contact CSV: $ContactCsvFile"

    if (-not (Get-Command gam -ErrorAction SilentlyContinue)) {
        throw "GAM コマンドが PATH で見つかりません"
    }

    if (-not (Test-Path -LiteralPath $ContactCsvFile)) {
        throw "連絡先CSVファイルが見つかりません: $ContactCsvFile"
    }

    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
    }

    Log-Message "--- Phase 1: 連絡先の登録処理開始 ---"

    $Contacts = Import-Csv -Path $ContactCsvFile -Encoding UTF8
    $count = 0
    $total = $Contacts.Count

    Log-Message "登録する連絡先数: $total 件"

    foreach ($Row in $Contacts) {
        $count++
        $fullName = "$($Row.LastName) $($Row.FirstName)"
        Log-Message "[$count / $total] 連絡先登録: $fullName"

        $NamesObj = @(@{
            givenName          = $Row.FirstName
            familyName         = $Row.LastName
            phoneticGivenName  = $Row.HiraganaFirstName
            phoneticFamilyName = $Row.HiraganaLastName
        })

        $ContactJson = @{
            names = $NamesObj
        }

        if (-not [string]::IsNullOrWhiteSpace($Row.OrganizationName) -or
            -not [string]::IsNullOrWhiteSpace($Row.OrganizationDepartment)) {
            $ContactJson.organizations = @(@{
                name       = $Row.OrganizationName
                department = $Row.OrganizationDepartment
                type       = "work"
            })
        }

        $EmailsObj = Get-EmailObjects -Row $Row
        if ($EmailsObj.Count -gt 0) {
            $ContactJson.emailAddresses = $EmailsObj
        }

        $PhonesObj = Get-PhoneObjects -Row $Row
        if ($PhonesObj.Count -gt 0) {
            $ContactJson.phoneNumbers = $PhonesObj
        }

        $ContactJson | ConvertTo-Json -Depth 10 | Out-File -FilePath $TempJsonFile -Encoding utf8

        $gamResult = Invoke-GamCommand -Arguments @(
            "user",
            $TargetUserEmail,
            "create",
            "contact",
            "json",
            "file",
            $TempJsonFile
        )

        if ($gamResult.ExitCode -ne 0) {
            Log-Message "    ! GAMコマンドエラー: $($gamResult.Output)"
        }
        elseif ($gamResult.Output) {
            Log-Message "    GAM出力: $($gamResult.Output)"
        }
    }

    Log-Message "--- Phase 1: 連絡先登録完了 ---"

    if (-not [string]::IsNullOrWhiteSpace($LabelCsvFile) -and (Test-Path -LiteralPath $LabelCsvFile)) {
        Log-Message "--- Phase 2: ラベルの一括作成開始 ---"
        Log-Message "Label CSV: $LabelCsvFile"

        try {
            $LabelColumns = Get-LabelColumns -CsvPath $LabelCsvFile

            if ($LabelColumns.Count -eq 0) {
                Log-Message "  警告: ラベル列が見つかりませんでした"
            }
            else {
                Log-Message "  検出されたラベル列: $($LabelColumns -join ', ')"

                foreach ($LabelCol in $LabelColumns) {
                    Log-Message "  処理中: $LabelCol"

                    $gamResult = Invoke-GamCommand -Arguments @(
                        "csv",
                        $LabelCsvFile,
                        "gam",
                        "user",
                        $TargetUserEmail,
                        "create",
                        "contactgroup",
                        "name",
                        "~$LabelCol"
                    )

                    if ($gamResult.Output) {
                        Log-Message "    GAM出力: $($gamResult.Output)"
                    }

                    if ($gamResult.ExitCode -ne 0) {
                        Log-Message "    GAMエラー ($LabelCol)"
                    }
                }

                Log-Message "  ラベル作成完了（$($LabelColumns.Count)列処理）"
            }
        }
        catch {
            Log-Message "  エラー: ラベル作成処理に失敗 - $_"
        }

        Log-Message "--- Phase 2: ラベル作成完了 ---"
    }
    else {
        Log-Message "--- Phase 2: ラベルCSV未指定のためスキップ ---"
    }

    Log-Message "=== 全ての処理が完了しました ==="
    exit 0
}
catch {
    Log-Message "重大なエラーが発生しました: $_"
    Log-Message "スタックトレース: $($_.ScriptStackTrace)"
    exit 1
}
finally {
    if (Test-Path -LiteralPath $TempJsonFile) {
        Remove-Item -LiteralPath $TempJsonFile -Force -ErrorAction SilentlyContinue
    }
}
