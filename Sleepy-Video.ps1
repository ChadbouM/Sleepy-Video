#region
################# 
# SLEEPY-VIDEO: #
##########################################################################
# A Utility for Watching Videos and then Falling Asleep!                 #
# Creates a context menu entry which plays videos while setting a        #
# timer to shutdown the computer one minute after the video ends.        #
# -                                                                      #
# Running this script installs/uninstalls the context menu entries.      #
# The menu entry relies on this script to then create the timed shutdown.#
# -----------------------------------------------------------------------#
# 1.0.0: Includes a second Menu Option, For Timing the video to end at   #
#        Specified Time.                                                 #
# Author: Michael Chadbourne                                             #
# Last Edited: 4/13/2016                                                 #
##########################################################################
#endregion

Param(
    [Parameter(Position=0)]
    [String] $InputPath, # Video Path for Run/Timed Modes

    [Switch] $Run,       # Switch for Watch Now Mode

    [Switch] $Timed      # Switch for Timed Mode
)

#region Constants
$CurrentVersion  = "1.0.0"
#
$ContextRegistry = "Registry::HKCR\*\shell"
$InstallRegistry = "Registry::HKLM\SOFTWARE\SleepyInstall"
$RunRegistry     = $ContextRegistry + '\SleepyVideo'
$TimedRegistry   = $ContextRegistry + '\SleepyTimed'
#
$MyName          = $myinvocation.MyCommand.Name
$DefaultDir      = "C:\Users\All Users"

$LengthConstant  = 27 #MP4 and MKV

#endregion

#region Functions:

#region Get-VideoLength( [String] ):
#    Return The String-Length of the Video at the given String-Directory
#endregion
Function Get-VideoLength([String]$VideoPath)
{
    $ParentFolder = Split-Path $VideoPath -Parent
    $FileName     = Split-Path $VideoPath -Leaf

    $FolderObject = New-Object -ComObject Shell.Application #Create Shell Object
    $FolderObject = $FolderObject.NameSpace($ParentFolder)  #Create Folder Object

    $FileObject   = $FolderObject.ParseName($FileName)      #Create FolderItem Object  

    Return $FolderObject.GetDetailsOf( $FileObject, $LengthConstant )
}

#region Play-SleepVideo( [String] ):
#    Plays the Video at the given location and
#    then shuts the computer down when the video has finished.
#endregion
Function Play-SleepVideo([String]$VideoPath)
{
    #Play The Video:
    . $VideoPath
    #Calculate Shutdown Time:
    $Time    = Get-Date
    $EndTime = Get-VideoLength($VideoPath)
    $EndDate = $Time + $EndTime + "00:01"
    #Write-Sleep-Shutdown
    Write-Output "Shutting Down at: $EndDate"
    Start-Sleep -Seconds ($EndDate - $Time).TotalSeconds
    Stop-Computer -Force
}

#region Play-TimedSleepVideo( [String] ):
#    Prompts the User to Enter a Time, then plays video at the
#    given address so that it will end at the given time. Then shuts
#    the computer off. If there is not enough time to finish the video
#    it will play instantly.
#endregion
Function Play-TimedSleepVideo([String]$VideoPath)
{
    # Initial Prompt
    Write-Output "When would you like to sleep?"
    Write-Output "Format: (hh:mm) ex. '22:30' for 10:30pm"
    # Read for Input and Calculate Times
    [datetime] $BedTime = Read-Host
    [datetime] $CurTime = Get-Date
    [timespan] $VidLeng = Get-VideoLength($VideoPath)
    # Handle Day-Loop
    If ($BedTime -lt $CurTime) {
        $BedTime = $BedTime.AddDays(1)
    }
    # Determining when to start the Video + Sleeping
    $StartTime = ($BedTime - $VidLeng - ([timespan] "00:01"))
    If ($StartTime -lt $CurTime) {
        Write-Output "Starting the Video Now"
        Write-Output "Shutting Down at $BedTime"
    } else {
        Write-Output "Starting video at: $StartTime"
        Write-Output "So it will end at: $BedTime"
        Start-Sleep -Seconds ($StartTime - $CurTime).TotalSeconds
    }
    #Play The Video:
    . $VideoPath
    $CurTime = Get-Date
    Start-Sleep -Seconds ($BedTime - $CurTime).TotalSeconds
    Stop-Computer -Force
}

#region Activate-Now( [String], [String] ):
#    Activates the 'Watch Now and Sleep' Context Menu Entry
#endregion
Function Activate-Now( [String] $InstallPath, [String] $ContextText )
{
    $CommandRegistry = $RunRegistry + '\command'
    $CommandValue    = "cmd /C PowerShell . '$InstallPath\SleepyVideo\Sleepy-Video.ps1' '%1' -Run"
    
    # Create the Registries
    New-Item -Path $RunRegistry     -Value $ContextText  | Out-Null
    New-Item -Path $CommandRegistry -Value $CommandValue | Out-Null
}

#region Activate-Timed( [String], [String] ):
#    Activates the 'Watch Before Bed' Context Menu Entry
#endregion
Function Activate-Timed( [String] $InstallPath, [String] $ContextText )
{
    $CommandRegistry = $TimedRegistry + '\command'
    $CommandValue    = "cmd /C PowerShell . '$InstallPath\SleepyVideo\Sleepy-Video.ps1' '%1' -Timed"
    
    # Create the Registries
    New-Item -Path $TimedRegistry   -Value $ContextText  | Out-Null
    New-Item -Path $CommandRegistry -Value $CommandValue | Out-Null
}

#region Deactivate-Now():
#    Deactivates the 'Watch Now and Sleep' Context Menu Entry
#endregion
Function Deactivate-Now()
{
    If (Test-Path -LiteralPath $RunRegistry) {
        # Remove the Registry
        Remove-Item -LiteralPath $RunRegistry -Recurse
    }
}

#region Deactivate-Timed():
#    Deactivates the 'Watch Before Bed' Context Menu Entry
#endregion
Function Deactivate-Timed()
{
    If (Test-Path -LiteralPath $TimedRegistry) {
        # Remove the Registry
        Remove-Item -LiteralPath $TimedRegistry -Recurse
    }
}

#region Detect-Version():
#    Scans the Registries to determine the latest version installed.
#    Returns the Version Numbers as a string, If No Installation is found
#    Then this Returns the String "-1"
#endregion
Function Detect-Version()
{
    #Look for 0.1.0
    If (Test-Path -LiteralPath $RunRegistry) {
        $ThisRegistry = Get-ItemProperty -LiteralPath $RunRegistry
        If ($ThisRegistry.Version -eq "0.1.0") { Return $ThisRegistry.Version }
    }
    #Look for Current Version (1.0.0)
    If ( Test-Path -LiteralPath $InstallRegistry ) {
        $ThisRegistry = Get-ItemProperty -LiteralPath $InstallRegistry
        If ($ThisRegistry.Version -ne $NULL) { Return $ThisRegistry.Version }
    }
    #No Known Version Found
    Return "-1"
}

#region Install-Script( [String] ):
#    Installs the Script, without activating Context Menu Items.
#    [ Creates an Installation Registry, and stores this Script
#      In the given Directory ]
#endregion
Function Install-Script( [String] $InstallPath )
{
    $FullPath     = $InstallPath     + '\SleepyVideo'    
    # Create the Install Registry
    New-Item         -Path        $InstallRegistry                                        | Out-Null
    New-ItemProperty -LiteralPath $InstallRegistry -Value $CurrentVersion -Name "Version" | Out-Null
    New-ItemProperty -LiteralPath $InstallRegistry -Value $InstallPath    -Name "Path"    | Out-Null
    # Create the Folder
    If (!(Test-Path -LiteralPath $FullPath)) { # If The Folder does not exist
        New-Item -Path $FullPath -ItemType directory | Out-Null
    }
    # Place this Script into the Folder
    Copy-Item -LiteralPath $PSCommandPath ($FullPath + '\Sleepy-Video.ps1')
}

#region Install-Update():
#    Updates the Registries and the Script in the Install Location.
#endregion
Function Install-Update()
{
    $InstalledVersion = Detect-Version
    # Upgrade from 0.1.0
    If ($InstalledVersion = "0.1.0") {
        # Identify Install Path
        $ThisRegistry  = Get-ItemProperty -LiteralPath $RunRegistry
        $InstallPath   = $ThisRegistry.Path
        $InstallText   = $ThisRegistry.'(default)'

        #Replace Script File
        If (Test-Path   -LiteralPath ($InstallPath + '\SleepyVideo\Sleepy-Video.ps1')) {
            Remove-Item -LiteralPath ($InstallPath + '\SleepyVideo\Sleepy-Video.ps1')
            } #TODO Standardize below
        Copy-Item -LiteralPath $PSCommandPath ($InstallPath + '\SleepyVideo\Sleepy-Video.ps1')

        #Replace Installation Registry
        Deactivate-Now
        $ThisRegistry = $ContextRegistry + '\SleepyInstall'   
        # Create the Install Registry
        New-Item         -Path        $InstallRegistry                                        | Out-Null
        New-ItemProperty -LiteralPath $InstallRegistry -Value $CurrentVersion -Name "Version" | Out-Null
        New-ItemProperty -LiteralPath $InstallRegistry -Value $InstallPath    -Name "Path"    | Out-Null

        #Activate 'Now' Mode
        Activate-Now $InstallPath $InstallText
        
        #Prompt Completion
        Write-Output "`nUpdate Complete"
    }
}

#region Install-Uninstall( [String] ):
#    Uninstalls the Context Menu Entries and removes
#    this script from the 'Install Path' which may or may not be 
#    this copy of the script.
#endregion
Function Install-Uninstall( [String] $InstallPath )
{
    #Deactivate Active Menu Items
        Deactivate-Now
        Deactivate-Timed
    #Remove Install Registry
    Remove-Item -LiteralPath $InstallRegistry -Recurse

    #Delete The Install Folder
    Remove-Item -LiteralPath ($InstallPath + '\SleepyVideo') -Recurse

    Write-Output "`nUninstall Complete"
}

#region Install-Automatic:
#    Detects the current Install state and prompts
#    the user for the appropriate actions
#endregion
Function Install-Automatic()
{
    # Detect Install:
    $Version = Detect-Version

    #Print Prompts based on Installation status:
    Clear 
    #If the Script is not installed
    If ($Version -eq "-1") { 
        Write-Output "Sleepy-Video is not currently Installed"
        Write-Output "Perform a Fresh Install? (Y/n)"
        $In = $host.UI.RawUI.ReadKey("IncludeKeyDown,NoEcho")
        If (!('n' -in $In.Character))
        {
            #Prompt for Install Location
            Write-Output "Where should it be installed?"
            Write-Output "Default: C:\Users\All Users\"
            Write-Output "(Just Hit Enter[Return] to use the Default)"

            $NewDir = Read-Host
            #Check for Default
            If ($NewDir -eq "") {
                $NewDir = $DefaultDir
            }
            Install-Script $NewDir
        }
    # If the Install is out of date
    } elseif ($Version -lt $CurrentVersion) { 
        Write-Output "Sleepy-Video is out of date!"
        Write-Output "Installed Version: $Version"
        Write-Output "Script Version:    $CurrentVersion"
        Write-Output "`nUpdate Now? (Y/n)"
        # Update if Not Prompted No
        $In = $host.UI.RawUI.ReadKey("IncludeKeyDown,NoEcho")
        If (!('n' -in $In.Character))
        {
            Install-Update
        }
    }
    # Update the Version
    $Version = Detect-Version
    #If the install is good!
    If ($Version -eq $CurrentVersion) {
        # Collect Information on the Install
        $InstallPath = (Get-ItemProperty -LiteralPath $InstallRegistry).Path
        # Check if NOW Mode is Active
        If (Test-Path -LiteralPath $RunRegistry) {
            $ActiveNow   = "Active"
        } else {
            $ActiveNow   = "Not Active"
        }
        # Check if Timed Mode is Active
        If (Test-Path -LiteralPath $TimedRegistry) {
            $ActiveTimed = "Active"
        } else {
            $ActiveTimed = "Not Active"
        }
        # Prompt Actions Available
        Write-Output "Sleepy-Video is Installed and up-to-date with this Script"
        Write-Output "Watch NOW:     $ActiveNow"
        Write-Output "Watch In Time: $ActiveTimed"
        Write-Output "`nHit (1) to Activate/Deactivate Watch-Now"
        Write-Output "Hit (2) to Activate/Deactivate Watch-in-Time"
        Write-Output "Hit (u) to Uninstall"
        Write-Output "Hit Any-Other-Key to Exit"

        # Read for User Choice
        $In = $host.UI.RawUI.ReadKey("IncludeKeyDown, NoEcho")
        # If they are Changing 'NOW'
        If ('1' -in $In.Character) {
            If ( $ActiveNow -eq "Not Active" ) {
                Write-Output "What should the Menu say?"
                Write-Output "Default: Watch Now Then Sleep"
                Write-Output "(Just Hit Enter[Return] to use the Default)"

                $ContextText = Read-Host
                #Check for Default
                If ($ContextText -eq "") {
                    $ContextText = "Watch Now Then Sleep"
                }
                Activate-Now $InstallPath $ContextText

            } else {
                Deactivate-Now
            }
        # If they are changing 'TIMED'
        } elseif ( '2' -in $In.Character ) {
            If ( $ActiveTimed -eq "Not Active" ) {
                Write-Output "`nWhat should the Menu say?"
                Write-Output "Default: Watch at Bedtime"
                Write-Output "(Just Hit Enter[Return] to use the Default)"

                $ContextText = Read-Host
                #Check for Default
                If ($ContextText -eq "") {
                    $ContextText = "Watch at Bedtime"
                }
                Activate-Timed $InstallPath $ContextText

            } else {
                Deactivate-Timed
            }
        # If they are uninstalling
        } elseif ( 'u' -in $In.Character ) {
            Install-Uninstall $InstallPath
        # If they are continuing
        } else {
            # All Done
        }
    }
}

#endregion

#region Script:
If ($Run) {                       #NOW MODE
    Play-SleepVideo $InputPath 
} elseif ($Timed) {               #TIMED MODE
    Play-TimedSleepVideo $InputPath 
} else {                          #INSTALLER MODE
    #Elevate to Administrator:
    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
    }
    #Run The Install Function
    Install-Automatic $InputPath 
}
#endregion