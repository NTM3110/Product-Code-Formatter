param(
    [string]$Version = "0.4.0",
    [string]$Notes = "Local development build",
    [ValidateSet("dev", "test", "stable")]
    [string]$Channel = "dev"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Không tìm thấy virtual environment: $python"
}

$oldBuildVersion = $env:PRODUCT_CODE_FORMATTER_BUILD_VERSION
$oldBuildNotes = $env:PRODUCT_CODE_FORMATTER_BUILD_NOTES
$oldBuildChannel = $env:PRODUCT_CODE_FORMATTER_BUILD_CHANNEL
$oldViteVersion = $env:VITE_APP_VERSION
try {
    $env:PRODUCT_CODE_FORMATTER_BUILD_VERSION = $Version
    $env:PRODUCT_CODE_FORMATTER_BUILD_NOTES = $Notes
    $env:PRODUCT_CODE_FORMATTER_BUILD_CHANNEL = $Channel
    $env:VITE_APP_VERSION = $Version

    Push-Location (Join-Path $root "react_frontend")
    try {
        npm run build
        if ($LASTEXITCODE -ne 0) { throw "React build thất bại." }
    } finally {
        Pop-Location
    }

    & $python -m py_compile (Join-Path $root "app.py") (Join-Path $root "web_api.py") (Join-Path $root "web_desktop_app.py") (Join-Path $root "product_code\update_client.py")
    if ($LASTEXITCODE -ne 0) { throw "Python compile thất bại." }

    Push-Location $root
    try {
        & $python -m PyInstaller ProductCodeFormatterWeb.spec --noconfirm
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller build thất bại." }
    } finally {
        Pop-Location
    }

    $appExe = Join-Path $root "dist\ProductCodeFormatter\ProductCodeFormatter.exe"
    if (-not (Test-Path -LiteralPath $appExe)) {
        throw "Không tìm thấy app sau build: $appExe"
    }
    Write-Output $appExe
} finally {
    $env:PRODUCT_CODE_FORMATTER_BUILD_VERSION = $oldBuildVersion
    $env:PRODUCT_CODE_FORMATTER_BUILD_NOTES = $oldBuildNotes
    $env:PRODUCT_CODE_FORMATTER_BUILD_CHANNEL = $oldBuildChannel
    $env:VITE_APP_VERSION = $oldViteVersion
}