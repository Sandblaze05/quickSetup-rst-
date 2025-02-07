# Check for Internet connection
function Test-InternetConnection {
    try {
        $ping=Test-Connection -TargetName "8.8.8.8" -Count 1 -Quiet -ErrorAction Stop
        if ($ping) {
            Write-Host "Internet connection is available."
            Check-Proxy
            return $true
        }
        else {
            Write-Host "Internet connection is not available."
            return $false
        }
    }
    catch {
        Write-Host "Internet connection is not available (an error has occured): $($_.Exception.Message)"
        return $false
    }
}

# Check for proxy settings
function Check-Proxy {
    $proxyEnabled = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -Name ProxyEnable -ErrorAction SilentlyContinue

    if ($proxyEnabled.ProxyEnable -eq 1) {
        $proxyServer = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -Name ProxyServer -ErrorAction SilentlyContinue
        Write-Host "Proxy is ENABLED: $($proxyServer.ProxyServer)"
        netsh winhttp import proxy source=ie | Out-Null # Import proxy settings for WinHTTP if proxy is enabled
        return 
    } else {
        Write-Host "Proxy is DISABLED."
        return
    }
}

# Check for Administrator privileges
function Ensure-Elevated {
    $adminCheck = [bool]((whoami /groups) -match "S-1-16-12288")
    if ($adminCheck) {
        Write-Host "Running as Administrator."
        return $true
    } else {
        Write-Host "Not running as Administrator."
        return $false
    }
}

#Check whether Chocolatey is installed
function Install-Chocolatey {
    if (-not (Test-Path "$env:ProgramData\chocolatey\choco.exe")) {
        Write-Host "Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
    else {
        Write-Host "Chocolatey detected."
    }
}

# Install a package using Chocolatey
function Install-Package {
    param (
        [string]$packageName
    )

    # Command name and choco package name mapping
    $hashMapping = @{
        "python" = "python3 --pre"
        "code" = "vscode"
        "gcc" = "mingw"
        "g++" = "mingw"
        "choco-cleaner" = "choco-cleaner"
    }

    if (-not (Test-Path "$env:ProgramData\chocolatey\choco.exe")) {
        Write-Host "Chocolatey is not installed. Aborting..."
        exit
    }
    if (-not (Already-Installed $packageName)) {
        Write-Host "Installing package $($hashMapping[$packageName])..."
        choco install $($hashMapping[$packageName]) -y
        if ($LASTEXITCODE -ne 0) {
            Write-Host "FAILED to install package $($hashMapping[$packageName])."
        }
        return
    }
}

# Check whether a package is already installed
function Already-Installed {
    param (
        [string]$packageName
    )

    if (Get-Command $packageName -ErrorAction SilentlyContinue) {
        Write-Host "Package $($hashMapping[$packageName]) is already installed."
        return $true
    }
    if (Get-Package -Name $packageName -ErrorAction SilentlyContinue) {
        Write-Host "Package $($hashMapping[$packageName]) is already installed."
        return $true
    }
    else {
        Write-Host "Package $($hashMapping[$packageName]) is not installed."
        return $false
    }
}

function Clean-Up {
    Write-Host "Cleaning up..."
    netsh winhttp reset proxy | Out-Null # Reset proxy settings for WinHTTP which were modified by Check-Proxy
    Install-Package "choco-cleaner"
    choco-cleaner
}

function main {
    if (-not (Ensure-Elevated)) {
        Write-Host "Please run this script as Administrator. Aborting..."
        exit
    }
    if (-not (Test-InternetConnection)) {
        Write-Host "Internet connection is not available. Aborting..."
        exit
    }
    Install-Chocolatey
    Install-Package "python"
    Clean-Up
}

main