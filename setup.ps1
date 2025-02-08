param (
    [switch]$logs,
    [switch]$packages,
    [switch]$office
)

# Check for Internet connection
function Test-InternetConnection {
    try {
        $ping=Test-Connection -TargetName "8.8.8.8" -Count 1 -Quiet -WarningAction Stop
        if ($ping) {
            Write-Host "Internet connection is available." -ForegroundColor Green
            Check-Proxy
            return $true
        }
        else {
            Write-Warning "Internet connection is not available."
            return $false
        }
    }
    catch {
        Write-Warning "Internet connection is not available (an Warning has occured): $($_.Exception.Message)"
        return $false
    }
}

# Check for proxy settings
function Check-Proxy {
    $proxyEnabled = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -Name ProxyEnable -WarningAction SilentlyContinue

    if ($proxyEnabled.ProxyEnable -eq 1) {
        $proxyServer = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -Name ProxyServer -WarningAction SilentlyContinue
        Write-Host "Proxy is ENABLED: $($proxyServer.ProxyServer)" -ForegroundColor Yellow
        netsh winhttp import proxy source=ie | Out-Null # Import proxy settings for WinHTTP if proxy is enabled
        return 
    } else {
        Write-Host "Proxy is DISABLED." -ForegroundColor Red
        return
    }
}

# Check for Administrator privileges
function Ensure-Elevated {
    $adminCheck = [bool]((whoami /groups) -match "S-1-16-12288")
    if ($adminCheck) {
        Write-Host "Running as Administrator." -ForegroundColor Green
        return $true
    } else {
        Write-Warning "Not running as Administrator."
        return $false
    }
}

# Check whether Chocolatey is installed
function Install-Chocolatey {
    if (-not (Test-Path "$env:ProgramData\chocolatey\choco.exe")) {
        Write-Host "Installing Chocolatey..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "FAILED to install Chocolatey."
            return
        }
        Write-Host "Chocolatey installed successfully." -ForegroundColor Green
        Write-Host "Refreshing environment variables..." -ForegroundColor Yellow
        Refresh-Environment
    }
    else {
        Write-Host "Chocolatey detected." -ForegroundColor Green
    }
}

# Declare the package mapping as a global variable
Set-Variable -Name packageMapping -Value @{
    "python" = "python3 --pre"
    "code" = "vscode"
    "gcc" = "mingw"
    "g++" = "mingw"
    "choco-cleaner" = "choco-cleaner"
    "java" = "openjdk"
    "node" = "nodejs"
} -Scope Global

# Install a package using Chocolatey
function Install-Package {
    param (
        [string]$packageName
    )

    if (-not (Test-Path "$env:ProgramData\chocolatey\choco.exe")) {
        Write-Warning "Chocolatey is not installed. Aborting..."
        exit
    }
    if (-not (Already-Installed $packageName)) {
        Write-Host "Installing package $($packageMapping[$packageName])..." -ForegroundColor Yellow
        choco install $($packageMapping[$packageName]) -y
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "FAILED to install package $($packageMapping[$packageName])."
        }
        return
    }
}

# Check whether a package is already installed
function Already-Installed {
    param (
        [string]$packageName
    )

    if (Get-Command $packageName -WarningAction SilentlyContinue) {
        Write-Host "Package $($packageMapping[$packageName]) is already installed." -ForegroundColor Green
        return $true
    }
    if (Get-Package -Name $packageName -WarningAction SilentlyContinue) {
        Write-Host "Package $($packageMapping[$packageName]) is already installed." -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "Package $($packageMapping[$packageName]) is not installed." -ForegroundColor Yellow
        return $false
    }
}

function Clean-Up {
    Write-Host "Cleaning up..." -ForegroundColor Yellow
    netsh winhttp reset proxy | Out-Null # Reset proxy settings for WinHTTP which were modified by Check-Proxy
    Install-Package "choco-cleaner"
    choco-cleaner
}

function Refresh-Environment {
    Write-Host "Refreshing environment variables..." -ForegroundColor Yellow
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
}


function main {
    param (
        [switch]$logs,
        [switch]$packages,
        [switch]$office
    )

    if ($packages) {
        Write-Host "`nPackages to be installed: $($packageMapping.Values)`n" -ForegroundColor Yellow
        exit
    }

    if ($logs) {
        Write-Host "Logging to setup.log..." -ForegroundColor Yellow
        Start-Transcript -Path "$env:TMP/setup.log" -Append | Out-Null
    }

    if (-not (Ensure-Elevated)) {
        Write-Warning "Please run this script as Administrator. Aborting..."
        if ($logs) { Stop-Transcript | Out-Null }
        exit
    }
    if (-not (Test-InternetConnection)) {
        Write-Warning "Internet connection is not available. Aborting..."
        if ($logs) { Stop-Transcript | Out-Null }
        exit
    }

    Install-Chocolatey
    Install-Package "python"
    Install-Package "g++"
    Install-Package "java"
    Install-Package "code"
    Install-Package "node"
    Clean-Up

    if ($office) {
        $officePaths = @(
            "HKLM:\SOFTWARE\Microsoft\Office",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
        )

        $officeInstalled = $false

        foreach ($path in $officePaths) {
            if (Test-Path $path) {
                Write-Host "Microsoft Office is installed." -ForegroundColor Green
                $officeInstalled = $true
                & ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook
                break
            }
        }

        if (-not $officeInstalled) {
            Write-Host "Microsoft Office is NOT installed." -ForegroundColor Red
        }

    }

    Write-Host "`nSetup completed successfully." -ForegroundColor Green

    if ($logs) { Stop-Transcript | Out-Null }
}

main -logs:$logs -packages:$packages -office:$office