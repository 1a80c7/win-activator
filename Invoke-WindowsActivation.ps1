<#
.SYNOPSIS
    Automated Windows activation utility for Windows 10/11 and Windows Server.

.DESCRIPTION
    The Invoke-WindowsActivation.ps1 script provides a silent, programmatic way to 
    activate Windows operating systems. It utilizes the Microsoft Activation Script 
    (MAS) HWID method for consumer editions (Windows 10/11) and the TSForge 
    method for Windows Server 2019 and newer versions.

.NOTES
    File Name      : Invoke-WindowsActivation.ps1
    Author         : Anthony Cotales
    Prerequisites  : Administrative privileges, Internet connectivity.
    Source         : 
#>


function Test-InternetConnection {
    <#
    .SYNOPSIS
        Tests for an active internet connection using a "No Content" web request.

    .DESCRIPTION
        This function checks connectivity by targeting Google's 'generate_204' endpoint.
        It uses -UseBasicParsing to bypass Internet Explorer dependencies, making it 
        ideal for server environments and fast execution.
    #>
    try {
        # Construct the URL for Google's 204 endpoint
        $url = ("https:", "", "www.google.com", "generate_204") -join "/"
        
        # Execute request with a 5-second timeout
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 5
        
        # Check the correct StatusCode property of the response object
        if ($response.StatusCode -eq 204) {
            return $true
        }
    }
    catch {
        # Handle DNS failures, timeouts, or 4xx/5xx errors
        return $false
    }
    return $false
}

function Invoke-ElevatedCheck {
    <#
    .SYNOPSIS
        Relaunches the current script with Administrator privileges.
    #>
    param(
        [switch]$NoExit
    )
    $params = @{
        FilePath     = "powershell.exe"
        Verb         = "RunAs"
        WindowStyle  = "Hidden"
        ArgumentList = (
            "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        )
    }

    if ($NoExit) {
        $params.ArgumentList = "-NoExit " + $params.ArgumentList
    }

    $admin = [Security.Principal.WindowsBuiltInRole]::Administrator
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal] $identity
 
    try {
        if (-not $principal.IsInRole($admin)) {
            Start-Process @params
            exit
        }
    }
    catch {
        # Handle when the operation is cancelled; Exit silently
        exit
    }
}

function Get-WindowsActivationStatus {
    <#
    .SYNOPSIS
        Checks Windows activation status using a specific WQL query.

    .DESCRIPTION
        Queries the SoftwareLicensingProduct class and returns a boolean 
        indicating whether the OS is fully licensed.
    #>
    try {
        # Define the query to find the Windows OS license info
        $query = @(
            "SELECT LicenseStatus FROM SoftwareLicensingProduct",
            "WHERE Name LIKE 'Windows%' AND PartialProductKey IS NOT NULL"
        ) -join " "

        # Execute the CIM instance search
        $licenseInfo = Get-CimInstance -Query $query -ErrorAction Stop

        # Check if the LicenseStatus property equals 1 (Licensed)
        # If multiple licenses are returned, .LicenseStatus becomes an array; 
        # -contains ensures we find at least one valid 'Licensed' status.
        $ActivationStatusBool = $licenseInfo.LicenseStatus -contains 1

        # Return a simple object with the boolean and the raw status
        return [PSCustomObject]@{
            IsActivated   = $ActivationStatusBool
            LicenseStatus = $licenseInfo.LicenseStatus
            ComputerName  = $env:COMPUTERNAME
        }
    }
    catch {
        Write-Error "Failed to check activation: $($_.Exception.Message)"
    }
}

function Get-RemoteFile {
    <#
    .SYNOPSIS
        Downloads a file from a URL and returns the local file path.

    .DESCRIPTION
        This function takes a source URL, downloads the file to a specified directory, 
        and returns the absolute string path. It includes error handling for network 
        and file system issues.

    .PARAMETER Url
        The direct download link for the file.

    .PARAMETER DestinationPath
        The folder where the file should be saved. Defaults to the current directory.
    #>
    param (
        [Parameter(Mandatory)]
        [string]$Url,
        [string]$DestinationPath = $PWD
    )

    # Save the current preference to restore it later
    $originalPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    try {
        # Validate the Destination Directory exists
        if (!(Test-Path -Path $DestinationPath)) {
            throw "The destination path '$DestinationPath' does not exist."
        }

        # Extract the filename and build the full path
        $fileName = Split-Path -Leaf $Url
        $fullPath = Join-Path -Path $DestinationPath -ChildPath $fileName

        # Attempt the download
        # -ErrorAction Stop ensures the catch block is triggered on HTTP errors
        Invoke-WebRequest -Uri $Url -OutFile $fullPath -ErrorAction Stop

        # Return the absolute path
        return (Get-Item $fullPath).FullName
    }
    catch {
        Write-Error "Download failed: $Url - $($_.Exception.Message)"
        return $null
    }
    finally {
        # Restore the original preference even if the download fails
        $ProgressPreference = $originalPreference
    }
}

function Get-WindowsProductType {
    <#
.SYNOPSIS
    Identifies if the current OS is a Workstation, Server, or Domain Controller.

.DESCRIPTION
    Queries CIM/WMI to retrieve the ProductType integer and maps it to a string.
#>
    try {
        # Retrieve the ProductType property
        $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        
        # Map the integer to a readable string
        switch ($osInfo.ProductType) {
            1 { return "Workstation" }
            2 { return "Domain Controller" }
            3 { return "Server" }
            Default { return "Unknown" }
        }
    }
    catch {
        Write-Error "Failed to detect OS type."
        return "Unknown"
    }
}

# --- BEGIN SCRIPT LOGIC ---
Invoke-ElevatedCheck

$URL = @(
    "https:", "", "github.com", "massgravel", "Microsoft-Activation-Scripts",
    "blob", "master", "MAS", "All-In-One-Version-KL", "MAS_AIO.cmd"
) -join "/"

$Activator = $null
$ActivatorParams = $null
$LicenseStatus = Get-WindowsActivationStatus
$ProductType = Get-WindowsProductType
    

if ($ProductType -eq "Workstation") {
    $ActivatorParams = '/HWID'
}
elseif ($ProductType -in ("Server", "Domain Controller")) {
    $ActivatorParams = '/Z-WindowsESUOffice'
}

if (Test-InternetConnection) {
    $Activator = Get-RemoteFile -Url $URL -DestinationPath $env:TEMP
}

if (Test-Path -Path $Activator -PathType Leaf -and $ActivatorParams) {
    $argsList = @('/c', $Activator, $ActivatorParams, '/S') -join " "
    Start-Process cmd.exe -ArgumentList $argsList -Verb RunAs -Wait
}

if ($LicenseStatus.IsActivated) {

    Add-Type -AssemblyName PresentationFramework

    [System.Windows.MessageBox]::Show(
        "Windows Activation Successful!", 
        "Microsoft Windows Activation", 
        [System.Windows.MessageBoxButton]::OK, 
        [System.Windows.MessageBoxImage]::Information
    )
}
