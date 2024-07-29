<#
.SYNOPSIS
This PowerShell script automates the process of downloading and extracting the latest versions of
tcpser a virtual modem emulator is required software to setup a virtual BBS development environment.

.DESCRIPTION
The script performs the following tasks:
1. Downloads the latest tcpser release from GitHub.
2. Fetches and downloads the latest Cygwin DLL from a specified mirror.
4. Extracts the downloaded files to specified directories.
6. Copies the Cygwin DLL to the tcpser directory.

.PARAMETER None

.EXAMPLE
.\install_tcpser.ps1

This example runs the script and starts the installation process by downloading the required software.

.NOTES
- The script sets the security protocol to TLS 1.2 to ensure secure connections.
- The script creates the download and extraction directories if they don't exist.
- Colorized messages are used to indicate the current process and ensure clarity during execution.
- Full installation of Cygwin is not required for tcpser usage.
#>


# Define the GitHub API URLs for the VICE releases and tcpser repository
$tcpserRepoUrl = "https://github.com/go4retro/tcpser/archive/refs/heads/master.zip"
$cygwinBaseUrl = "https://mirror.steadfast.net/cygwin/x86_64/release/cygwin/"

# Define the paths for download and extraction
$downloadPath = "$env:TEMP"
$extractPath = "C:\tools"


# Check if 7-Zip is installed
$7zipPath = "C:\Program Files\7-Zip\7z.exe"  # Adjust this path if 7-Zip is installed elsewhere
if (-Not (Test-Path -Path $7zipPath)) {
    Write-Host "7-Zip is not installed. Downloading and installing..." -ForegroundColor Red

    # Define the URL for the 7-Zip installer and the output path
    $installerUrl = "https://www.7-zip.org/a/7z1900-x64.exe"  # Update URL to the latest version if necessary
    $installerPath = "$env:TEMP\7zInstaller.exe"

    # Download the installer
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

    # Execute the installer
    # Note: /S for silent install, adjust arguments as necessary
    $process = Start-Process -FilePath $installerPath -Args "/S" -Wait -NoNewWindow -PassThru

    # Check the exit code of the process
    if ($process.ExitCode -ne 0) {
	Write-Host "The installation process exited with an error code: $($process.ExitCode)" -ForegroundColor Red
	exit 1
    } else {
 	Write-Host "7-Zip installation complete." -ForegroundColor Green
    }
} else {
    Write-Host "7-Zip is already installed." -ForegroundColor Green
}

# Function to extract tar.xz files
function Expand-7Ziptarxz {
    param (
        [string]$FilePath,
        [string]$ExtractPath
    )

    $7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
    if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
        Write-Host "7-Zip executable '$7zipPath' not found."
        exit 1
    }

    try {
        # make sure this is the xz file
       	if ($FilePath -match '.xz$') {
            #unpack xz with 7zip		
            $xzUnpack = & $7zipPath x $FilePath -o"$downloadPath" *>&1
            if ($xzUnpack) {
                $tarFilename = $FilePath -replace '\.xz$', ''
                #make sure tar file
                if ($tarFilename -match '.tar$') {
                    # Extract tar
                    $tarExtract = & $7zipPath x $tarFilename -o"$ExtractPath" *>&1
                    if ($tarExtract) {
                        Remove-Item $tarFilename
                        return $true
                    } else {
                        Throw "Failed to extract $tarFilename"
                    }
                }
            } else {
                Throw "Failed to unpack $xzUnpack"
            }
        } else {
            Throw "Failed to find correct filename $xzFilename"
        }
    } catch {
        Write-Host "Error extracting archive: $_"
        return $false
    }
}

# Function to set the security protocol to TLS 1.2
function Set-SecurityProtocol {
    param (
        [Net.SecurityProtocolType]$Protocol = [Net.SecurityProtocolType]::Tls12
    )
    [Net.ServicePointManager]::SecurityProtocol = $Protocol
}

# Function to create a new directory if it does not exist
function New-DirectoryExists {
    param (
        [string]$Path
    )
    if (-Not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory | Out-Null
        Write-Host "Created directory: $Path"
    }
}

# Ensure the download and extract directories exist
New-DirectoryExists -Path $downloadPath
New-DirectoryExists -Path $extractPath

# Set the security protocol
Set-SecurityProtocol -Protocol Tls12

# Display the banner
Write-Host "Starting Installation - Downloading Requirements" -ForegroundColor White

# Download tcpser
Write-Host "Checking for existing tcpser download..." -ForegroundColor Yellow
$tcpserOutputPath = "$downloadPath\tcpser.zip"
if (-Not (Test-Path -Path $tcpserOutputPath)) {
    Write-Host "Downloading tcpser..." -ForegroundColor Blue
	try {
        Invoke-WebRequest -Uri $tcpserRepoUrl -OutFile $tcpserOutputPath
        Write-Host "tcpser download complete." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download tcpser. Error: $_" -ForegroundColor Red
	exit 1
    }
} else {
    Write-Host "tcpser is already downloaded." -ForegroundColor Green
}

# Get the latest Cygwin DLL version
Write-Host "Fetching latest Cygwin DLL version..." -ForegroundColor Cyan
# Fetch the Cygwin page
$cygwinPage = Invoke-WebRequest -Uri $cygwinBaseUrl -UseBasicParsing
# Extract the latest version link
$cygwinLatestVersion = ($cygwinPage.Links | Where-Object { $_.href -match "cygwin-.*\.tar\.xz" -and $_.href -notmatch "src" } | Sort-Object href -Descending | Select-Object -First 1).href
# Construct the download URL
$cygwinDllUrl = "$cygwinBaseUrl$cygwinLatestVersion"
# Set output path and file name
$cygwinDllArchivePath = "$downloadPath\$($cygwinLatestVersion -replace '.*/')"
# Download
Write-Host "Checking for existing tcpser download..." -ForegroundColor Yellow
# Check if the file already exists
if (-Not (Test-Path -Path $cygwinDllArchivePath)) {
    Write-Host "Downloading Cygwin" -ForegroundColor Blue
    try {
        # Download the Cygwin DLL
        Invoke-WebRequest -Uri $cygwinDllUrl -OutFile $cygwinDllArchivePath
        Write-Host "Cygwin DLL download complete." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download Cygwin. Error: $_" -ForegroundColor Red
	exit 1
    }
} else {
    Write-Host "Cygwin is already downloaded." -ForegroundColor Green
}

# Display completion message
Write-Host "Downloads Complete - Ready to Extract" -ForegroundColor White

# Extract tcpser
Write-Host "Checking for extracted tcpser dir..." -ForegroundColor Yellow
$tcpserExtractPath = "$extractPath\tcpser"
if (-Not (Test-Path -Path "$tcpserExtractPath\tcpser.exe")) {
    Write-Host "Extracting tcpser..." -ForegroundColor Blue
    Try {
        Expand-Archive -Path $tcpserOutputPath -DestinationPath $extractPath -Force
        # Rename the directory
        $sourceDir = "$extractPath\tcpser-master"
        if (Test-Path -Path $sourceDir) {
            Rename-Item -Path $sourceDir -NewName "tcpser"
            Write-Host "tcpser extraction complete." -ForegroundColor Green
        } else {
            Write-Error "Failed to find the directory in the archive."
            exit 1
        }
        Write-Host "tcpser extraction complete." -ForegroundColor Green
    } Catch {
        Write-Error "Failed to expand the archive: $_"
        exit 1
    }
} else {
    Write-Host "tcpser already extracted." -ForegroundColor Green
}

# Extract Cygwin
$cygwinExtractPath = "$downloadPath\cygwin"
Write-Host "Checking for cygwin extract folder..." -ForegroundColor Yellow
# Check if extracted folder exists
if (Test-Path $cygwinExtractPath) {
    # look for cygwin dir in temp
    Write-Host "Cygwin extract path found." -ForegroundColor Green
    if (-Not (Test-Path -Path "$cygwinExtractPath\usr\bin\cygwin1.dll")) {
        Throw "Extracted file: cygwin1.dll not found."
    }
    Write-Host "cygwin1.dll extract file found." -ForegroundColor Green
} else {
    #extracted folder cygwin does not exist
    Write-Host "Extracting cygwin1.dll..." -ForegroundColor Blue
    if (Test-Path -Path $cygwinDllArchivePath) {
        #looking for download
        Write-Host "Cygwin download exists." -ForegroundColor Red
	    #tar -xaf $cygwinDllArchivePath -C $cygwinExtractPath
	    $tarxzout = & Expand-7Ziptarxz -FilePath $cygwinDLLArchivePath -ExtractPath $cygwinExtractPath 
        if ( $tarxzout ) {
            Write-Host "Extraction successful!"
        } else {
            Throw "Extraction failed."
        }
        if (Test-Path -Path "$cygwinExtractPath\usr\bin\cygwin1.dll") {
            Write-Host "cygwin1.dll extraction complete." -ForegroundColor Green
        } else {
       	    Throw "Extraction failed: cygwin1.dll not found."
        }
    }
}

# Copy cygwin1.dll to the tcpser folder if it does not exist
Write-Host "Checking for cygwin1.dll in tcpser folder..." -ForegroundColor Yellow
$destinationDllPath = "$tcpserExtractPath\cygwin1.dll"
if (-Not (Test-Path -Path $destinationDllPath)) {
    Write-Host "cygwin1.dll not found in tcpser. Checking for source dll..." -ForegroundColor Blue
    $cygwinDllPath = "$cygwinExtractPath\usr\bin\cygwin1.dll"
    if (Test-Path -Path $cygwinDllPath) {
        Write-Host "Copying cygwin1.dll to tcpser folder..." -ForegroundColor Blue
	    Try {
            Copy-Item -Path $cygwinDllPath -Destination $destinationDllPath -Force
            Write-Host "cygwin1.dll copy complete." -ForegroundColor Green
        } Catch {
            Write-Error "Failed to copy cygwin1.dll: $_"
            exit 1
        }
    } else {
        Write-Host "Source cygwin1.dll not found. Copy operation aborted." -ForegroundColor Red
    }
} else {
    Write-Host "cygwin1.dll already exists in tcpser folder." -ForegroundColor Green
}

# List of expected downloaded files
$downloadedFiles = @(
    $tcpserOutputPath,
	$destinationDllPath
    # Add other paths as necessary
)

# Check if all downloads are completed
$allDownloadsCompleted = $true
foreach ($file in $downloadedFiles) {
    if (-Not (Test-Path -Path $file)) {
        Write-Host "Download incomplete or missing for file: $file" -ForegroundColor Red
        $allDownloadsCompleted = $false
    }
}

# Check if all downloads are completed test if working
if ($allDownloadsCompleted) {
    Write-Host "Download and Extracted successfully." -ForegroundColor White

    # Define the path to tcpser.exe
    $tcpserPath = "$tcpserExtractPath\tcpser.exe"

    # Run the command and capture the output
    $output = & $tcpserPath /? 2>&1 | Out-String

    # Check if the command executed successfully
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Command failed to execute."
        exit 1
    }

    # Check if the output contains the word "Usage:"
    if ($output -match "Usage:") {
        Write-Host "Command executed successfully and returned usage information."
    } else {
        Write-Host "Command did not return the expected usage information."
        exit 1
    }

	# FUN
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(932,150)
	Start-Sleep -m 75
	[System.Console]::Beep(1047,150)
	Start-Sleep -m 75
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[console]::Beep(699,150)
	Start-Sleep -m 75
	[System.Console]::Beep(740,150)
	Start-Sleep -m 75
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(932,150)
	Start-Sleep -m 75
	[System.Console]::Beep(1047,150)
	Start-Sleep -m 75
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(784,150)
	Start-Sleep -m 150
	[System.Console]::Beep(699,150)
	Start-Sleep -m 75
	[System.Console]::Beep(740,150)
	Start-Sleep -m 75
	[System.Console]::Beep(932,150)
	[System.Console]::Beep(784,150)
	[System.Console]::Beep(587,1200)
	Start-Sleep -m 36
	[System.Console]::Beep(932,150)
	[System.Console]::Beep(784,150)
	[System.Console]::Beep(554,1200)
	Start-Sleep -m 36
	[System.Console]::Beep(932,150)
	[System.Console]::Beep(784,150)
	[System.Console]::Beep(523,1200)
	Start-Sleep -m 75
	[System.Console]::Beep(466,150)
	[System.Console]::Beep(523,150)
	exit 0 # success
} else {
    Write-Host "One or more downloads are incomplete or missing." -ForegroundColor Red
	exit 1 # success
}
