param(
    [string]$InputFile = "C:\Users\TT1\Desktop\RPA抓取hot不全数据.xlsx",
    [string]$OutputDir = "",
    [ValidateSet("chrome", "edge")]
    [string]$Browser = "chrome"
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $scriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "未找到项目虚拟环境。请先按 README 执行环境初始化。"
}

$arguments = @(
    (Join-Path $scriptRoot "tanji_hot_contacts.py"),
    "--input", $InputFile,
    "--config", (Join-Path $scriptRoot "config.json")
)
$arguments += @("--browser", $Browser)
if ($OutputDir) {
    $arguments += @("--output-dir", $OutputDir)
}

& $pythonExe @arguments
exit $LASTEXITCODE
