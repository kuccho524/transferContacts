param(
    [Parameter(Mandatory = $true)]
    [string]$GamArgsJson
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

try {
    if (-not (Get-Command gam -ErrorAction SilentlyContinue)) {
        throw "GAM コマンドが PATH で見つかりません"
    }

    $GamArgs = @($GamArgsJson | ConvertFrom-Json)
    $result = & gam @GamArgs 2>&1
    $exitCode = $LASTEXITCODE

    if ($null -ne $result) {
        $result | Out-String -Width 4096 | Write-Output
    }

    exit $exitCode
}
catch {
    Write-Error $_
    exit 1
}
