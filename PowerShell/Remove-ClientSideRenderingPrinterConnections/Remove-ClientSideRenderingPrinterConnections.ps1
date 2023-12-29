<#
.SYNOPSIS
    Removes cached client side rendering printer connections from the registry
    to fix 'Group Policy Printers' Event ID 4098 error code '0x80070057'.

.VERSION
    1.0.0, 29/12/2023

.AUTHOR
    George Escott https://github.com/georgeescott

.LINK
    https://github.com/georgeescott/Scripts/tree/master/PowerShell/Remove-ClientSideRenderingPrinterConnections

.DESCRIPTION
    This script fixes the 'Group Policy Printers' Event ID 4098 error:
    "The user '<printer name>' preference item in the '<GPO name>
    {00000000-0000-0000-0000-000000000000}' Group Policy Object did not apply
    because it failed with error code: '0x80070057 The parameter is incorrect.'
    This error was suppressed.".

.NOTES
    This script should be run when no users are logged in. Ideally at either
    startup, shutdown or as a scheduled task.

    This script should be used in combination with the following registry key
    (which this script sets by default):

    [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider]
    "RemovePrintersAtLogoff"=dword:00000001

    See the GitHub README for more information.

#>

################################################################################
################################## FUNCTIONS ###################################
################################################################################

<#
.SYNOPSIS
    Returns the current timestamp. For the purposes of outputting a timestamp to
    a log file.
#>
function timestamp {
    return "[" + $(Get-Date -format "yyyy-MM-dd HH:mm:ss") + "]"
}

<#
.SYNOPSIS
    Creates or sets the RemovePrintersAtLogoff registry key to 1
#>
function Set-RemovePrintersAtLogoff {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        if ( $null -eq (Get-ItemProperty -Path $CSRPrintProviderPath -Name 'RemovePrintersAtLogoff' -ErrorAction SilentlyContinue)."RemovePrintersAtLogoff" ){
            # Create RemovePrintersAtLogoff with a value of 1 if it doesn't exist
            New-ItemProperty -Path $CSRPrintProviderPath -Name 'RemovePrintersAtLogoff' -Value 1 -PropertyType DWORD -Force
            Write-Output "$(timestamp) Registry key '$CSRPrintProviderPath\RemovePrintersAtLogoff' created with a value of 1."

        } elseif ( (Get-ItemProperty -Path $CSRPrintProviderPath -Name 'RemovePrintersAtLogoff' -ErrorAction SilentlyContinue)."RemovePrintersAtLogoff" -ne 1 ){
            # Update RemovePrintersAtLogoff to a value of 1 if it exists but is the incorrect value
            Set-ItemProperty -Path $CSRPrintProviderPath -Name 'RemovePrintersAtLogoff' -Value 1
            Write-Output "$(timestamp) Registry key '$CSRPrintProviderPath\RemovePrintersAtLogoff' updated to a value of 1."

        } else {
            Write-Output "$(timestamp) Registry key '$CSRPrintProviderPath\RemovePrintersAtLogoff' is already set to 1."
        }
    } catch {
        Write-Error "$(timestamp) An error occurred while setting the '${CSRPrintProviderPath}\RemovePrintersAtLogoff' registry key: $_"
    }
}


<#
.SYNOPSIS
    Removes any cached domain user SID's from the 'Client Side Rendering Print
    Provider' registry key.

.DESCRIPTION
    When 'RemovePrintersAtLogoff' is set, the following sub-keys are removed:
    HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\S-1-5-21-*

    This function will replicate this functionality and remove these keys for
    cached user profiles.
#>
function Remove-CSRPrintProviderSIDs {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        if ( Test-Path $CSRPrintProviderPath ){
            # SID Structures: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dtyp/81d92bba-d22b-4a8c-908a-554ab29148ab
            $SIDRegex = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'

            # Get a list of cached domain user SID's in the 'Client Side Rendering Print Provider' key
            $CSRPrintProviderSIDs = (Get-Item $CSRPrintProviderPath).GetSubKeyNames() -match $SIDRegex

            # If cached SID's exist, remove them
            # This removes sub-keys matching 'S-1-5-21-*' under the 'Client Side Rendering Print Provider' key
            if ( $CSRPrintProviderSIDs.Count -gt 0){
                Write-Output "$(timestamp) Found cached user SID's under '${CSRPrintProviderPath}'"

                $CSRPrintProviderSIDs | ForEach-Object {
                    $SIDkey = "$CSRPrintProviderPath\$_"

                    if ( (Test-Path $SIDkey) -and ($SIDkey -ne "$CSRPrintProviderPath\") ){
                        Write-Output "$(timestamp) Removing '${SIDkey}'"
                        Remove-Item –path "$SIDkey" -Recurse -Force
                    }
                }
            } else {
                Write-Output "$(timestamp) No cached user SID's found under '${CSRPrintProviderPath}'"
            }
        }
    } catch {
        Write-Error "$(timestamp) An error occurred while removing SID's from the '${CSRPrintProviderPath}' registry key: $_"
    }
}


<#
.SYNOPSIS
    Removes any cached Printers and Monitors from the 'Client Side Rendering
    Print Provider' Servers registry key.

.DESCRIPTION
    When 'RemovePrintersAtLogoff' is set, the following sub-keys are removed:
    HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers\<name>\Printers\*
    HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers\<name>\Monitors\Client Side Port\*

    This function will replicate this functionality and remove these keys for
    cached user profiles.
#>
function Remove-CSRPrintProviderServers {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        if ( Test-Path "$CSRPrintProviderPath\Servers" ){
            $CSRPrintProviderServers = (Get-Item "$CSRPrintProviderPath\Servers").GetSubKeyNames()

            # If servers were returned, loop through each server and remove the Printers and Monitor sub-keys
            if ( $CSRPrintProviderServers.Count -gt 0){
                Write-Output "$(timestamp) Found print server(s) under '${CSRPrintProviderPath}\Servers'"

                $CSRPrintProviderServers | ForEach-Object {
                    $Server = "$CSRPrintProviderPath\Servers\$_"

                    if ( (Test-Path $Server) -and ($Server -ne "$CSRPrintProviderPath\Servers\") ){
                        $ServerPrintersPath = "$Server\Printers"

                        if ( (Get-Item $ServerPrintersPath -ErrorAction SilentlyContinue).SubKeyCount -ge 1 ){
                            Write-Output "$(timestamp) Found cached Printer keys under '${ServerPrintersPath}'"
                            Write-Output "$(timestamp) Removing '$ServerPrintersPath\*'"
                            Remove-Item -Path "$ServerPrintersPath\*" -Recurse -Force
                        }

                        $ServerMonitorsClientSidePortPath = "$Server\Monitors\Client Side Port"

                        if ( (Get-Item $ServerMonitorsClientSidePortPath -ErrorAction SilentlyContinue).SubKeyCount -ge 1 ){
                            Write-Output "$(timestamp) Found cached Client Side Port keys under '${ServerMonitorsClientSidePortPath}'"
                            Write-Output "$(timestamp) Removing '$ServerMonitorsClientSidePortPath\*'"
                            Remove-Item -Path "$ServerMonitorsClientSidePortPath\*" -Recurse -Force
                        }
                    }
                }
            } else {
                Write-Output "$(timestamp) No print servers found under '${$CSRPrintProviderPath}\Servers'"
            }
        } else {
            Write-Output "$(timestamp) '${CSRPrintProviderPath}\Servers' does not exist"
        }
    } catch {
        Write-Error "$(timestamp) An error occurred while removing cached Printers and Monitors from the '${CSRPrintProviderPath}\Servers' registry key: $_"
    }
}


<#
.SYNOPSIS
    Restarts the Print Spooler service
#>
function Restart-PrintSpooler {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        $ServiceName = 'Spooler'
        $ServiceDisplayName = 'Printer Spooler'

        Write-Output "$(timestamp) Restarting '${ServiceDisplayName}' service..."
        Restart-Service -Name $ServiceName -Force

        # Wait for 5 seconds
        Start-Sleep 5

        # Contigency check to make sure it's running
        $PrintSpooler = Get-Service $ServiceName

        if ( $PrintSpooler.Status -ne 'Running'){
            Write-Output "$(timestamp) '${ServiceDisplayName}' service not running after 5 seconds."

            Write-Output "$(timestamp) Attempting to start '${ServiceDisplayName}' service..."
            $PrintSpooler.Start()

            # Wait for the service to reach the $status for a maximum of 10 seconds
            Write-Output "$(timestamp) Waiting a maximum of 10 seconds for it to start..."
            $PrintSpooler.WaitForStatus("Running", '00:00:10')
        } elseif ( $PrintSpooler.Status -eq 'Running' ){
            Write-Output "$(timestamp) '${ServiceDisplayName}' service restarted."
        }
    } catch {
        Write-Error "$(timestamp) An error occurred while restarting the '${ServiceDisplayName}' service: $_"
    }
}


################################################################################
##################################### MAIN #####################################
################################################################################

# Start the Transcript for logging
Start-Transcript -Path "$env:WINDIR\Logs\$($MyInvocation.MyCommand.Name).log" -Append

# Set the registry path for the Client Side Rendering Print Provider
$CSRPrintProviderPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider'

# Check that the 'Client Side Rendering Print Provider' key exists
if ( Test-Path $CSRPrintProviderPath ){
    Set-RemovePrintersAtLogoff

    Remove-CSRPrintProviderSIDs

    Remove-CSRPrintProviderServers

    #Restart-PrintSpooler
} else {
    try {
        # Create the 'Client Side Rendering Print Provider' key if it doesn't exist
        Write-Output "$(timestamp) '${CSRPrintProviderPath}' does not exist. Creating key."
        New-Item -Path $CSRPrintProviderPath -Force | Out-Null
    } catch {
        Write-Error "$(timestamp) An error occurred while setting the '${CSRPrintProviderPath}' registry key: $_"
    }
}

Write-Output "$(timestamp) Script Complete."

# End transcript
Stop-Transcript