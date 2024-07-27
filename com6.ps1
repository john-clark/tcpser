<#
.SYNOPSIS
This PowerShell script is a wrapper for the tcpser a virtual modem emulator which is required 
to setup a virtual BBS development environment.

*** THIS SCRIPT IS NOT FINISHES AND NEEDS TO BE TESTED ***

.DESCRIPTION
The script performs the following tasks:
1. Detects if a IP232 port or a COM port is specified.
2. If a COM port initializes the specified COM port for use.
3. Starts the tcpser.exe process with the specified options.

.EXAMPLE
Example - start TCPSER with comport COM6 and TCP port 6400
> Start-Tcpser $portName $baudRate $logLevel $tcpserOptions $tcpserIncomingPort

Example - start TCPSER with virtual IP232 port 25232 and TCP port 6400
> Start-Tcpser $tcpserip232Port $baudRate $logLevel $tcpserOptions $tcpserIncomingPort

Example - stop TCPSER process
> stop-Tcpser com6

Example - Install telnet client to test the connection
> dism /online /Enable-Feature /FeatureName:TelnetClient
> telnet localhost 25232

.NOTES
- The script checks if the specified COM port exists and initializes it if necessary.
- This script is intended to be used in conjunction with the tcpser virtual modem emulator.
- The script logs messages to a specified log file for troubleshooting and tracking.
- The script can be modified to include additional options for the TCPSER process.
- Currently, the script only supports starting and stopping _ONE_ TCPSER process.

#>

# $tcpserPath: Path to the TCPSER executable
$tcpserPath = "c:\tools\tcpser\tcpser.exe"

# $tcpserOptions: Additional options for TCPSER
# -v <port>: Specify a virtual IP232 port instead of a physical serial port.
$tcpserip232Port = 25232
# -d $portName: Specifies the COM port to use (e.g., COM5)
$portName = "COM6"
# -s $baudRate: Sets the baud rate for the COM port (e.g., 38400)
$baudRate = 38400
# -l $logLevel: Sets the logging level (e.g., 7 for detailed logging)
$logLevel = 7
# -tsSiI: 
#   -t: Enables tracing of the modem commands and responses.
#   -s: Enables server mode, allowing incoming connections.
#   -S: Enables server mode with secure (SSL/TLS) connections.
#   -i: Enables initialization string processing.
# -i 's0=1': Sets the modem to auto-answer on the first ring.
$tcpserOptions = "-tsSiI -i 's0=1'"
# -p $tcpserPort: Specifies the TCP port for TCPSER to listen on (e.g., 6400)
$tcpserIncomingPort = 6400
# -c <file>: Send the contents of a file to the local serial connection upon connect.
# -C <file>: Send the contents of a file to the remote IP connection upon connect.
# -a <file>: Send the contents of a file to the local serial connection upon answer.
# -A <file>: Send the contents of a file to the remote IP connection upon answer.
# -B <file>: Send the contents of a file upon detecting a busy signal.

# -I <file>: Send the contents of a file upon no answer. 
# -T <timeout>: Set an inactivity timeout in seconds.


# Registry path for serial port parameters
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Serial\Parameters"
$portKey = "PortName"

# Function to log messages with timestamp and type
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Type
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Write-Output $logEntry
    Add-Content -Path "C:\logs\tcpser\tcpser.log" -Value $logEntry
}

# Function to check if the COM port exists
function Test-COMPortExists {
    param (
        [string]$portName
    )

    try {
        $comPorts = Get-WmiObject Win32_SerialPort | Select-Object -ExpandProperty DeviceID
        return $comPorts -contains $portName
    } catch {
        Write-Error "Error checking COM port: $_"
        Write-LogMessage "Error checking COM port: $_" "ERROR"
        return $false
    }
}

# Function to initialize a serial port for use
function Initialize-SerialPort {
    param (
        [string]$portName
    )

    if (Test-COMPortExists -portName $portName) {
        Write-LogMessage "$portName already exists." "INFO"
    } else {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force
        }
        New-ItemProperty -Path $regPath -Name $portKey -Value $portName -PropertyType String -Force
        Get-PnpDevice -PresentOnly | Where-Object { $_.Status -eq 'Error' } | ForEach-Object { $_ | Enable-PnpDevice -Confirm:$false }
        Write-LogMessage "$portName has been added and hardware changes rescanned." "INFO"

        # Restart the serial service
        Restart-Service -Name "Serial" -Force
        # Wait for the service to restart
        Start-Sleep -Seconds 5
        # Check if the COM port exists after restart
        if (Test-COMPortExists -portName $portName) {
            Write-LogMessage "$portName has been initialized and the Serial service has been restarted." "INFO"
        } else {
            Write-LogMessage "Failed to initialize $portName and restart the Serial service." "ERROR"
            return $false
        }
    }
    return $true
}

# Wrapper Function to start tcpser.exe with the specified options
function Start-Tcpser {
    param (
        [ValidateScript({ Test-Path $_ })] [string]$TcpserPath,
        $checkHostingPort,
        [ValidateRange(300, 115200)] [int]$BaudRate,
        [ValidateRange(0, 7)] [int]$LogLevel,
        [string]$TcpserOptions,
        [ValidateRange(1, 65535)] [int]$tcpserIncomingPort
    )

    # Check if TCPSER path exists
    if (-not (Test-Path $TcpserPath)) {
        Write-Error "The specified path '$TcpserPath' does not exist."
        Write-LogMessage "The specified path '$TcpserPath' does not exist." "ERROR"
        return
    }

    # Check if TCPSER is already running
    $process = Get-Process -Name "tcpser.exe" -ErrorAction SilentlyContinue
    if ($process) {
        Write-Output "TCPSER is already running."
        Write-LogMessage "TCPSER is already running." "INFO"
        return
    }

    # Check if the COM port path is a port or a path
    if ($checkHostingPort -match '^COM\d+$') {
        $portName = $checkHostingPort
        # Check if the COM port exists
        if (-not (Initialize-SerialPort -portName $portName)) {
            return
        }
        $servPortOption = "-d $portName"
    } else {
        # Handle other cases, e.g., IP:port or integer port
        if ($checkHostingPort -match '^\d+$' -or $checkHostingPort -match '^\d{1,3}(\.\d{1,3}){3}:\d+$') {
            $servPortOption = "-v $checkHostingPort"
        } else {
            throw "Invalid port format: $checkHostingPort"
        }
    }


    # Make sure $servPortOption is set
    if (-not $servPortOption) {
        throw "servPortOption is not set"
    }

    # Start TCPSER with the specified options
    try {
        $arguments = "$servPortOption -s $BaudRate -l $LogLevel $TcpserOptions -p $tcpserIncomingPort"
        Start-Process -FilePath $TcpserPath -ArgumentList $arguments
        Write-LogMessage "TCPSER started successfully." "INFO"
    } catch {
        Write-Error "Error starting TCPSER: $_"
        Write-LogMessage "Error starting TCPSER: $_" "ERROR"
    }
}

# Wrapper function to stop TCPSER
function Stop-Tcpser {
    param (
        [string]$ProcessName = "tcpser.exe"
    )

    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($process) {
        Stop-Process -Name $ProcessName -Force
        Write-Output "Process $ProcessName has been forcefully terminated."
    } else {
        Write-Output "Process $ProcessName is not running."
    }
}

# Example - start TCPSER with comport COM6 and TCP port 6400
Start-Tcpser $portName $baudRate $logLevel $tcpserOptions $tcpserIncomingPort
