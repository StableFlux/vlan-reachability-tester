# VLAN Reachability Tester - one-command build script
# Usage: powershell -ExecutionPolicy Bypass -File build.ps1
# Optionally pass a version: build.ps1 -Version 1.2.0.0

param(
    [string]$Version = "1.0.0.0",
    [switch]$Demo
)

$repoDir    = $PSScriptRoot
$srcDir     = "$repoDir\Windows"
$demoDir    = "$srcDir\demo"
$packageDir = "$repoDir\VLANPackage"
$assetsDir  = "$packageDir\Assets"
$output     = "$repoDir\VLANReachabilityTester_$($Version)_x64.msix"
$demoExe    = "$demoDir\VLANReachabilityTester_Demo.exe"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
$label = if ($Demo) { "Build $Version (DEMO)" } else { "Build $Version" }
Write-Host "  VLAN Reachability Tester - $label" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($Demo) {
    foreach ($f in @("vlan_config.demo.json", "vlan_results.demo.json")) {
        if (-not (Test-Path "$demoDir\$f")) {
            Write-Host "ERROR: Demo seed file missing: $demoDir\$f" -ForegroundColor Red
            exit 1
        }
    }
}

# Step 1: PyInstaller
Write-Host "[1/3] Building exe with PyInstaller..." -ForegroundColor Yellow

# Remove any existing exe up front. If it's locked (a prior instance is still
# running) PyInstaller silently falls back to the stale exe and the rest of
# the build packages outdated code without warning.
$exe = "$srcDir\VLANReachabilityTester.exe"
if (Test-Path $exe) {
    Remove-Item $exe -Force -ErrorAction SilentlyContinue
    if (Test-Path $exe) {
        Write-Host "ERROR: Cannot remove $exe - a running instance is holding it open." -ForegroundColor Red
        Write-Host "       Close all VLANReachabilityTester.exe processes and re-run the build." -ForegroundColor Red
        exit 1
    }
}

Push-Location $srcDir
& pyinstaller --onefile --windowed --name "VLANReachabilityTester" vlan_tester_gui.py `
    --distpath "$srcDir" --workpath "$srcDir\build" --specpath "$srcDir" 2>&1 |
    Select-Object -Last 5
Pop-Location

# Clean up PyInstaller build artifacts
Remove-Item "$srcDir\build" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$srcDir\VLANReachabilityTester.spec" -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $exe)) {
    Write-Host "ERROR: PyInstaller failed - exe not found." -ForegroundColor Red
    exit 1
}
Write-Host "  Exe built: $exe" -ForegroundColor Green

# Demo short-circuit: stage a self-contained portable exe next to the seed
# JSONs and stop. No MSIX is produced - the portable.flag sentinel tells the
# app to read/write data in this folder instead of %LOCALAPPDATA%, so the demo
# never touches production data.
if ($Demo) {
    Write-Host "[2/2] Staging portable demo folder..." -ForegroundColor Yellow

    if (Test-Path $demoExe) { Remove-Item $demoExe -Force }
    Copy-Item $exe $demoExe -Force

    # Portable-mode sentinel (empty file is sufficient; its presence is the signal).
    [System.IO.File]::WriteAllText("$demoDir\portable.flag", "", [System.Text.UTF8Encoding]::new($false))

    # Runtime copies of the seeds - the app reads these on launch. Overwritten
    # by the app as it pings; re-run `build.ps1 -Demo` to restore.
    Copy-Item "$demoDir\vlan_config.demo.json"  "$demoDir\vlan_config.json"  -Force
    Copy-Item "$demoDir\vlan_results.demo.json" "$demoDir\vlan_results.json" -Force

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  DEMO BUILD SUCCESS!" -ForegroundColor Green
    Write-Host "  $demoDir" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Double-click VLANReachabilityTester_Demo.exe - fully self-contained." -ForegroundColor Cyan
    Write-Host ""
    Start-Process explorer.exe $demoDir
    exit 0
}

# Step 2: Package layout
Write-Host "[2/3] Creating package layout..." -ForegroundColor Yellow

Remove-Item $packageDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null

Copy-Item $exe $packageDir -Force

# Resize logo to required asset sizes
Add-Type -AssemblyName System.Drawing
$logo = [System.Drawing.Image]::FromFile("$repoDir\images\logo.png")

function Resize-Image($img, $width, $height, $path) {
    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($img, 0, 0, $width, $height)
    $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

Resize-Image $logo 44  44  "$assetsDir\Square44x44Logo.png"
Resize-Image $logo 150 150 "$assetsDir\Square150x150Logo.png"
Resize-Image $logo 310 150 "$assetsDir\Wide310x150Logo.png"
Resize-Image $logo 50  50  "$assetsDir\StoreLogo.png"
Resize-Image $logo 620 300 "$assetsDir\SplashScreen.png"
$logo.Dispose()

# Write manifest
[System.IO.File]::WriteAllText("$packageDir\AppxManifest.xml", @"
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
         xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
         xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
         IgnorableNamespaces="uap rescap">
  <Identity Name="StableFlux.VLANReachabilityTester"
            Publisher="CN=157477FA-4F07-4A30-A07E-126CF85871A8"
            Version="$Version"
            ProcessorArchitecture="x64"/>
  <Properties>
    <DisplayName>VLAN Reachability Tester</DisplayName>
    <PublisherDisplayName>StableFlux</PublisherDisplayName>
    <Logo>Assets\StoreLogo.png</Logo>
  </Properties>
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.26100.0"/>
  </Dependencies>
  <Resources>
    <Resource Language="en-us"/>
  </Resources>
  <Applications>
    <Application Id="VLANReachabilityTester" Executable="VLANReachabilityTester.exe" EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements DisplayName="VLAN Reachability Tester"
                          Description="Test network reachability across VLANs in real time"
                          BackgroundColor="transparent"
                          Square150x150Logo="Assets\Square150x150Logo.png"
                          Square44x44Logo="Assets\Square44x44Logo.png">
        <uap:DefaultTile Wide310x150Logo="Assets\Wide310x150Logo.png"/>
        <uap:SplashScreen Image="Assets\SplashScreen.png"/>
      </uap:VisualElements>
    </Application>
  </Applications>
  <Capabilities>
    <rescap:Capability Name="runFullTrust"/>
  </Capabilities>
</Package>
"@, [System.Text.UTF8Encoding]::new($false))

# Step 3: Pack MSIX
Write-Host "[3/3] Packing MSIX..." -ForegroundColor Yellow

$makeappx = Get-ChildItem "C:\Program Files (x86)\Microsoft Visual Studio\Shared\NuGetPackages\microsoft.windows.sdk.buildtools" -Recurse -Filter "makeappx.exe" -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -match "x64" } | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName

if (-not $makeappx) {
    $makeappx = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin" -Recurse -Filter "makeappx.exe" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match "x64" } | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
}

if (-not $makeappx) {
    Write-Host "ERROR: makeappx.exe not found." -ForegroundColor Red
    exit 1
}

if (Test-Path $output) { Remove-Item $output -Force }
& $makeappx pack /d $packageDir /p $output /nv | Out-Null

# Clean up the package staging folder
Remove-Item $packageDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
if (Test-Path $output) {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  SUCCESS!" -ForegroundColor Green
    Write-Host "  $output" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next: upload this file to Partner Center:" -ForegroundColor Cyan
    Write-Host "  https://partner.microsoft.com/dashboard" -ForegroundColor Cyan
    Write-Host ""
    Start-Process explorer.exe (Split-Path $output)
} else {
    Write-Host "FAILED: MSIX was not created." -ForegroundColor Red
    exit 1
}
