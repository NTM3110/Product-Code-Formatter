param(
    [switch]$ResetData,
    [Alias("Port")]
    [int]$ApiPort = 8765,
    [int]$FrontendPort = 5173
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $root ".venv/Scripts/python.exe"
$frontend = Join-Path $root "react_frontend"
$localAppData = [Environment]::GetFolderPath("LocalApplicationData")
$runtimeRoot = Join-Path $localAppData "ProductCodeFormatterPreview"
$runtimeLocalAppData = Join-Path $runtimeRoot "LocalAppData"
$runtimeAppData = Join-Path $runtimeLocalAppData "ProductCodeFormatter"
$runtimeWorkspace = Join-Path $runtimeRoot "workspace"
$runtimeViteCache = Join-Path $runtimeRoot "vite-cache"
$installedAppData = Join-Path $localAppData "ProductCodeFormatter"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Python environment not found: $python"
}
if (-not (Test-Path -LiteralPath (Join-Path $frontend "node_modules"))) {
    throw "node_modules is missing. Run npm install in react_frontend first."
}
$node = (Get-Command node.exe -ErrorAction Stop).Source
$viteCli = Join-Path $frontend "node_modules/vite/bin/vite.js"
if (-not (Test-Path -LiteralPath $viteCli)) {
    throw "Vite CLI not found: $viteCli"
}

if ($ResetData -and (Test-Path -LiteralPath $runtimeRoot)) {
    $separator = [IO.Path]::DirectorySeparatorChar
    $resolvedRoot = [IO.Path]::GetFullPath($localAppData).TrimEnd($separator)
    $resolvedRuntime = [IO.Path]::GetFullPath($runtimeRoot).TrimEnd($separator)
    if (-not $resolvedRuntime.StartsWith($resolvedRoot + $separator, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to delete runtime outside LocalAppData: $resolvedRuntime"
    }
    Remove-Item -LiteralPath $runtimeRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $runtimeAppData -Force | Out-Null
New-Item -ItemType Directory -Path $runtimeWorkspace -Force | Out-Null
New-Item -ItemType Directory -Path $runtimeViteCache -Force | Out-Null

$runtimeConfig = Join-Path $runtimeAppData "product_code_config.json"
if (-not (Test-Path -LiteralPath $runtimeConfig)) {
    foreach ($name in @("product_code_config.json", "license.json", "default_form_mappings.json")) {
        $source = Join-Path $installedAppData $name
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination (Join-Path $runtimeAppData $name) -Force
        }
    }
    $installedTemplates = Join-Path $installedAppData "templates"
    if (Test-Path -LiteralPath $installedTemplates) {
        Copy-Item -LiteralPath $installedTemplates -Destination $runtimeAppData -Recurse -Force
    }
    if (-not (Test-Path -LiteralPath $runtimeConfig)) {
        $repoConfig = Join-Path $root "product_code_config.json"
        if (Test-Path -LiteralPath $repoConfig) {
            Copy-Item -LiteralPath $repoConfig -Destination $runtimeConfig -Force
        }
    }
}

function Assert-PortAvailable([int]$Port, [string]$Label) {
    $listener = $null
    try {
        $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
        $listener.Start()
    } catch {
        throw "$Label port $Port is already in use. Close the old preview or choose another port."
    } finally {
        if ($null -ne $listener) {
            $listener.Stop()
        }
    }
}

function Wait-Vite([Diagnostics.Process]$Process, [string]$Url, [string]$ErrorLog) {
    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        $Process.Refresh()
        if ($Process.HasExited) {
            $detail = if (Test-Path -LiteralPath $ErrorLog) {
                Get-Content -LiteralPath $ErrorLog -Raw
            } else {
                "No log is available."
            }
            throw "Vite stopped before the UI opened: $detail"
        }
        try {
            Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 2 | Out-Null
            return
        } catch {
            Start-Sleep -Milliseconds 400
        }
    }
    throw "Vite did not start within 45 seconds. See $ErrorLog"
}

Assert-PortAvailable $ApiPort "FastAPI"
Assert-PortAvailable $FrontendPort "Vite"

$oldLocalAppData = $env:LOCALAPPDATA
$oldRuntimeDir = $env:PRODUCT_CODE_FORMATTER_RUNTIME_DIR
$oldConfigPath = $env:PRODUCT_CODE_FORMATTER_CONFIG_PATH
$oldTitle = $env:PRODUCT_CODE_FORMATTER_APP_TITLE
$oldPort = $env:PRODUCT_CODE_FORMATTER_PORT
$oldFrontendUrl = $env:PRODUCT_CODE_FORMATTER_FRONTEND_URL
$oldViteVersion = $env:VITE_APP_VERSION
$oldViteApiUrl = $env:VITE_API_URL
$oldViteCacheDir = $env:VITE_CACHE_DIR
$viteProcess = $null

try {
    $env:LOCALAPPDATA = $runtimeLocalAppData
    $env:PRODUCT_CODE_FORMATTER_RUNTIME_DIR = $runtimeWorkspace
    $env:PRODUCT_CODE_FORMATTER_CONFIG_PATH = $runtimeConfig
    $env:PRODUCT_CODE_FORMATTER_APP_TITLE = "ProductCodeFormatter - Runtime Preview"
    $env:PRODUCT_CODE_FORMATTER_PORT = [string]$ApiPort
    $env:PRODUCT_CODE_FORMATTER_FRONTEND_URL = "http://127.0.0.1:$FrontendPort"
    $env:VITE_APP_VERSION = "PREVIEW"
    $env:VITE_API_URL = "http://127.0.0.1:$ApiPort"
    $env:VITE_CACHE_DIR = $runtimeViteCache

    & $python -m py_compile (Join-Path $root "app.py") (Join-Path $root "web_api.py") (Join-Path $root "web_desktop_app.py")
    if ($LASTEXITCODE -ne 0) { throw "Python compile failed." }

    $viteOut = Join-Path $runtimeRoot "vite.stdout.log"
    $viteError = Join-Path $runtimeRoot "vite.stderr.log"
    Remove-Item -LiteralPath $viteOut, $viteError -Force -ErrorAction SilentlyContinue
    $viteArgs = '"' + $viteCli + '" --host 127.0.0.1 --port ' + [string]$FrontendPort + ' --strictPort'
    $viteProcess = Start-Process -FilePath $node -ArgumentList $viteArgs -WorkingDirectory $frontend -PassThru -WindowStyle Hidden -RedirectStandardOutput $viteOut -RedirectStandardError $viteError
    Wait-Vite $viteProcess $env:PRODUCT_CODE_FORMATTER_FRONTEND_URL $viteError

    Write-Host ""
    Write-Host "Running Runtime Preview from source..." -ForegroundColor Cyan
    Write-Host "UI: $env:PRODUCT_CODE_FORMATTER_FRONTEND_URL"
    Write-Host "API: http://127.0.0.1:$ApiPort"
    Write-Host "Preview data: $runtimeAppData"
    Write-Host "Close the app window to stop the runtime."
    & $python (Join-Path $root "web_desktop_app.py")
    if ($LASTEXITCODE -ne 0) { throw "Runtime Preview exited with code $LASTEXITCODE." }
} finally {
    if ($null -ne $viteProcess) {
        $viteProcess.Refresh()
        if (-not $viteProcess.HasExited) {
            Stop-Process -Id $viteProcess.Id -Force -ErrorAction SilentlyContinue
        }
    }
    $env:LOCALAPPDATA = $oldLocalAppData
    $env:PRODUCT_CODE_FORMATTER_RUNTIME_DIR = $oldRuntimeDir
    $env:PRODUCT_CODE_FORMATTER_CONFIG_PATH = $oldConfigPath
    $env:PRODUCT_CODE_FORMATTER_APP_TITLE = $oldTitle
    $env:PRODUCT_CODE_FORMATTER_PORT = $oldPort
    $env:PRODUCT_CODE_FORMATTER_FRONTEND_URL = $oldFrontendUrl
    $env:VITE_APP_VERSION = $oldViteVersion
    $env:VITE_API_URL = $oldViteApiUrl
    $env:VITE_CACHE_DIR = $oldViteCacheDir
}
