param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$Notes = "ProductCodeFormatter $Version",
    [ValidateSet("test", "stable")]
    [string]$Channel = "test",
    [switch]$SkipRemoteSync,
    [switch]$BundleOnly
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$feedPath = "updates/product-code-formatter"
$releaseDir = Join-Path (Join-Path $root "Releases") $Channel
$bundleStage = Join-Path (Join-Path (Join-Path $root "build") "release_bundle") (Join-Path $Channel $Version)
$deployDir = Join-Path $root "deploy"
$packId = if ($Channel -eq "stable") { "ProductCodeFormatter.App" } else { "ProductCodeFormatter.App.Test" }
$packTitle = if ($Channel -eq "stable") { "ProductCodeFormatter" } else { "ProductCodeFormatter Test" }
$existingFeed = Join-Path $releaseDir ("releases." + $Channel + ".json")

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
New-Item -ItemType Directory -Path $deployDir -Force | Out-Null

if (-not $BundleOnly) {
    if (-not (Get-Command vpk -ErrorAction SilentlyContinue)) {
        throw "Không tìm thấy Velopack CLI. Cài một lần bằng: dotnet tool install --global vpk --version 1.2.0"
    }

    if (-not $SkipRemoteSync) {
        $servers = @("http://192.168.1.210:8080", "http://192.168.101.13:8080")
        foreach ($server in $servers) {
            vpk download http --url ($server + "/" + $feedPath) --channel $Channel --outputDir $releaseDir --timeout 1
            if ($LASTEXITCODE -eq 0) { break }
            Write-Warning "Không đồng bộ được release cũ từ $server; thử server tiếp theo."
        }
    }

    if (Test-Path -LiteralPath $existingFeed) {
        try {
            $feed = Get-Content -LiteralPath $existingFeed -Raw -Encoding UTF8 | ConvertFrom-Json
            $requestedBase = [version](($Version -split '[-+]')[0])
            $blockingRelease = $feed.Assets | Where-Object {
                [version](($_.Version -split '[-+]')[0]) -ge $requestedBase
            } | Select-Object -First 1
            if ($blockingRelease) {
                throw "Channel $Channel đã có version $($blockingRelease.Version), không thể build $Version. Hãy tăng version."
            }
        } catch [System.Management.Automation.RuntimeException] {
            throw
        } catch {
            throw "Không đọc được feed hiện tại $existingFeed : $($_.Exception.Message)"
        }
    }

    & (Join-Path $root "build_app.ps1") -Version $Version -Notes $Notes -Channel $Channel
    if ($LASTEXITCODE -ne 0) { throw "Build app thất bại." }

    $notesDir = Join-Path (Join-Path $root "build") "release_notes"
    New-Item -ItemType Directory -Path $notesDir -Force | Out-Null
    $notesPath = Join-Path $notesDir ($Channel + "-" + $Version + ".md")
    [IO.File]::WriteAllText($notesPath, $Notes, [Text.UTF8Encoding]::new($false))

    vpk pack --packId $packId --packVersion $Version --packDir (Join-Path (Join-Path $root "dist") "ProductCodeFormatter") --mainExe "ProductCodeFormatter.exe" --packTitle $packTitle --packAuthors "ProductCodeFormatter" --channel $Channel --outputDir $releaseDir --releaseNotes $notesPath --icon (Join-Path $root "app_icon.ico") --shortcuts "Desktop,StartMenuRoot"
    if ($LASTEXITCODE -ne 0) { throw "Velopack pack thất bại." }
}

if (-not (Test-Path -LiteralPath $existingFeed)) {
    throw "Không tìm thấy feed $existingFeed. Hãy build một release trước."
}

$feed = Get-Content -LiteralPath $existingFeed -Raw -Encoding UTF8 | ConvertFrom-Json
$feedPackageNames = @($feed.Assets | ForEach-Object { Split-Path -Leaf ([string]$_.FileName) })
$currentPackage = $feed.Assets | Where-Object {
    [string]$_.Version -eq $Version -and [string]$_.Type -eq "Full"
} | Select-Object -First 1
if (-not $currentPackage) {
    throw "Feed $Channel không có full package version $Version."
}

if (Test-Path -LiteralPath $bundleStage) {
    Remove-Item -LiteralPath $bundleStage -Recurse -Force
}
New-Item -ItemType Directory -Path $bundleStage -Force | Out-Null

$releaseFiles = Get-ChildItem -LiteralPath $releaseDir -File
$releaseMetadataNames = @("releases.$Channel.json", "assets.$Channel.json", "RELEASES-$Channel")
$setupNameInFeed = "$packId-$Channel-Setup.exe"
$selected = $releaseFiles | Where-Object {
    $_.Name -in $releaseMetadataNames -or
    $_.Name -in $feedPackageNames -or
    $_.Name -eq $setupNameInFeed
}
if (-not ($selected | Where-Object { $_.Name -eq "releases.$Channel.json" })) {
    throw "Release feed releases.$Channel.json không được tạo."
}
foreach ($packageName in $feedPackageNames) {
    if (-not ($selected | Where-Object { $_.Name -eq $packageName })) {
        throw "Feed tham chiếu package chưa có trong Releases: $packageName"
    }
}
if (-not ($selected | Where-Object { $_.Name -eq $setupNameInFeed })) {
    throw "Setup package cho channel $Channel không được tạo."
}
$selected | Copy-Item -Destination $bundleStage -Force

$bundlePath = Join-Path $deployDir ("ProductCodeFormatter_" + $Channel + "_v" + $Version + "_bundle.zip")
if (Test-Path -LiteralPath $bundlePath) { Remove-Item -LiteralPath $bundlePath -Force }
Compress-Archive -Path (Join-Path $bundleStage "*") -DestinationPath $bundlePath -CompressionLevel Optimal

$setup = Join-Path $releaseDir $setupNameInFeed
$setupName = if ($Channel -eq "stable") { "ProductCodeFormatter-Setup.exe" } else { "ProductCodeFormatter-Test-Setup.exe" }
$versionedSetupName = if ($Channel -eq "stable") { "ProductCodeFormatter-Setup-v$Version.exe" } else { "ProductCodeFormatter-Test-Setup-v$Version.exe" }
Copy-Item -LiteralPath $setup -Destination (Join-Path $deployDir $setupName) -Force
Copy-Item -LiteralPath $setup -Destination (Join-Path $deployDir $versionedSetupName) -Force

Write-Output "Release bundle: $bundlePath"
Write-Output "Velopack output: $releaseDir"