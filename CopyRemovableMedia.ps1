[CmdletBinding(HelpURI="https://github.com/fpdavis/CopyRemovableMedia", SupportsShouldProcess)]
Param (
    [switch]$Help,
    [switch]$UseNewestFolder,     
    [switch]$SubfolderSearch, 
    [switch]$CopyToPopulatedDirectory,     
    [switch]$Quiet, 
    [switch]$NoEject, 
    [switch]$ExitAfterCopy, 
    [switch]$NoAutoUpdate, 
    $Source = '', 
    $Destination = '.', 
    [goVerbosityEnum]$Verbosity = [goVerbosityEnum]::Verbose
    )

$Version = 1.2

Add-Type -TypeDefinition @"
    public enum goVerbosityEnum {
    Critical,
    Error,
    Warning,
    Information,
    Verbose,
    Debug
}
"@

Function Main {
                                                                                                                                                                                                                                                                                                           
    if ($Help) {
        DisplayHelp
    }

    if ($Verbosity -eq "Debug") {
        $DebugPreference = "Continue"
     }
     else {
        $DebugPreference = "SilentlyContinue"
     }

     CheckForUpdate

     MessageLog "`nInsert media to be copied... Ctrl-C to cancel"

    while ($true) {    
        $CDROMDrive = Get-WMIObject -Class Win32_CDROMDrive

        if ($CDROMDrive.MediaLoaded) {
            CopyProcess

            if (-Not $ExitAfterCopy) {
                MessageLog "`nInsert media to be copied... Ctrl-C to cancel"
            }
        }

        start-sleep (5); 
    }
}

Function DisplayHelp {

    $Hash = CalculateHash $MyInvocation.ScriptName

    MessageLog ("`nNAME")
    MessageLog ("`tCopyRemovableMedia.ps1 (Version " + $Version + ", Hash: " + $Hash.Hash + ")")
    
	MessageLog "`nSYNTAX"
    MessageLog ("`t" + $MyInvocation.ScriptName.Replace((Split-Path $MyInvocation.ScriptName),'').TrimStart('\') + "`n")
    MessageLog "`t`t[-Help]"
    MessageLog "`t`t[-UseNewestFolder]"
    MessageLog "`t`t[-SubfolderSearch]"
    MessageLog "`t`t[-CopyToPopulatedDirectory]"    
    MessageLog "`t`t[-Quiet]"
    MessageLog "`t`t[-NoEject]"    
    MessageLog "`t`t[-ExitAfterCopy]"    
    MessageLog "`t`t[-NoAutoUpdate]"      
    MessageLog "`t`t[-Confirm]"  
    MessageLog "`t`t[-WhatIf]"  
    MessageLog "`t`t[-Source 'c:\Your\Source\Directory']"
    MessageLog "`t`t[-Destination 'c:\Your\Destination\Directory']"    
    MessageLog "`t`t[-Verbose]"
	
    MessageLog "`nDESCRIPTION"
    MessageLog "`t`tThis program was designed to facilitate the manual process of copying and archiving removable"
    MessageLog "`t`tmedia such as CDROM, DVD, and Blu-ray drives. With no parameters specified the program will identify"
    MessageLog "`t`tthe computers Optical Disk Drive and if there is a disk in the drive it will retrieve the name of"
    MessageLog "`t`tthe volume and create a new directory with the same name in the current directory. It will then"
    MessageLog "`t`tproceed to copy the contents of the entire disk into the newly created directory. A chime is played"
    MessageLog "`t`twhen the copy is complete to prompt the user to insert another disk for copy."
        
    MessageLog "`nPARAMETERS"
    MessageLog "`t`t                    Help - Displays this help screen alond with removable/optical drive information"
    MessageLog "`t`t`t                         and then exits."
    MessageLog "`t`t         UseNewestFolder - By default the program will create a new folder in the Destination"
    MessageLog "`t`t`t                         folder with the Volume Name of the drive being copied from."
    MessageLog "`t`t`t                         This parameter bypasses the new folder creation and searches the"
    MessageLog "`t`t`t                         destination folder for the folder with the newest creation date."
    MessageLog "`t`t`t                         This folder is then used as the final destination for the file"
    MessageLog "`t`t`t                         copy. Subfolders will not be searched by default unless the"
    MessageLog "`t`t`t                         SubfolderSearch parameter is specified. An existing folder with"
    MessageLog "`t`t`t                         files wills in it will not be written to unless the"
    MessageLog "`t`t`t                         CopyToPopulatedDirectory parameter is specified."
    MessageLog "`t`t         SubfolderSearch - Searches the subfolders of the Destination directory when looking for the"
    MessageLog "`t`t`t                         newest folder. Otherwise the newest folder in the root directory will"
    MessageLog "`t`t`t                         be choosen."
    MessageLog "`t`tCopyToPopulatedDirectory - By default the program will only copy files to an empty folder."
    MessageLog "`t`t`t                         By specifying this paramater you are giving the program permission"
    MessageLog "`t`t`t                         to copy files into an existing directory that already contains files."
    MessageLog "`t`t`t                         This paramater should be used sparingly and with caution."
    MessageLog "`t`t                   Quiet - Refrains from playing a chime when the copy is complete."
    MessageLog "`t`t                   Eject - Refrains from ejecting the optical drive after a copy is complete."    
    MessageLog "`t`t           ExitAfterCopy - If this parameter is not specified the program will do multiple copies"
    MessageLog "`t`t`t                         one after another. If specified the program will exit after the first"
    MessageLog "`t`t`t                         copy is made."
    MessageLog "`t`t            NoAutoUpdate - Skips checking for an update online."
    MessageLog "`t`t                 Confirm - The SupportsShouldProcess argument adds Confirm and WhatIf parameters to"
    MessageLog "`t`t`t                         the function. The Confirm parameter prompts the user before it performs"
    MessageLog "`t`t`t                         a copy (and before it creates the destination directory if necessary)."
    MessageLog "`t`t                  WhatIf - The SupportsShouldProcess argument adds Confirm and WhatIf parameters to"
    MessageLog "`t`t`t                         the function. The WhatIf parameter lists the changes that the command"
    MessageLog "`t`t`t                         would make, instead of running the commands."
    MessageLog "`t`t                  Source - By default the program identifies the optical drive to copy from. This can"
    MessageLog "`t`t`t                         be overriden."
    MessageLog "`t`t`t                         with a different drive and even a complete file path."
    MessageLog "`t`t             Destination - This is the current directory by default. This is the directory where a"
    MessageLog "`t`t`t                         new folder"
    MessageLog "`t`t`t                         will be created based on the Voume Name of the Source Disk. If the"
    MessageLog "`t`t`t                         UseNewestFolder is specified then this is the directory that will be"
    MessageLog "`t`t`t                         searched for the newest empty folder to copy the source contents into."
    MessageLog "`t`t               Verbosity - Displays various levels of messaging."
    MessageLog "`t`t`t                         Debug - Super detail"
    MessageLog "`t`t`t                         Verbose – Everything Important"
    MessageLog "`t`t`t                         Information - Updates that might be useful to user"
    MessageLog "`t`t`t                         Warning - Bad things that happen, but are expected"
    MessageLog "`t`t`t                         Error - Expected errors that are not recoverable"
    MessageLog "`t`t`t                         Critical - Unexpected errors that stop execution"
    
    MessageLog "`nEXAMPLES"
    MessageLog "`t`t"

    MessageLog "`nNOTES"
    MessageLog "`t`tDoes not recurse into symbolicly linked directores."
    MessageLog "`t`tLimited support for USB and floppy drives. These drive letters must be explicitly specified as the Source."
    
    # Drive type uses a numerical code:
    #
    # 0 -- Unknown
    # 1 -- No Root directory
    # 2 -- Removable Disk
    # 3 -- Local Disk
    # 4 -- Network Drive
    # 5 -- Compact Disc
    # 6 -- Ram Disk
    exit
    MessageLog "`nREMOVABLE(2)/OPTICAL(5) DRIVE INFORMATION"    
    Get-CimInstance Win32_LogicalDisk | ?{ $_.DriveType -eq 2 -or $_.DriveType -eq 5} 
    
    exit
}

Function CopyProcess {

    $DirectoryCreationTime = [DateTime]::MinValue
    $DestinationDirectory = ""
    $VolumeName = ""

    if ($Source -eq "") {    
        MessageLog ("Detecting optical drive.`n") ([goVerbosityEnum]::Verbose)

        $CDROMDrive = Get-WMIObject -Class Win32_CDROMDrive

        if (-Not $CDROMDrive) {
            MessageLog ("No CD/DVD/BR drive found. Copy aborted.`n") ([goVerbosityEnum]::Error)
            return
        }

        $Source = $CDROMDrive.Drive + "\"
        if (-Not $CDROMDrive.MediaLoaded) {
            MessageLog ("No disk in drive (" + $Source + "). Copy aborted.`n") ([goVerbosityEnum]::Error)
            return
        }

        $VolumeName = $CDROMDrive.VolumeName    
    }
    else {    
        MessageLog ("Checking soure drive (" + $Source.Substring(0,2) + ").`n") ([goVerbosityEnum]::Verbose)

        $DiskInfo = Get-CimInstance Win32_LogicalDisk | ?{ $_.DeviceID -eq $Source.Substring(0,2) -and ($_.DriveType -eq 5 -Or $_.DriveType -eq 2)} 
    
        if (-Not $DiskInfo.VolumeName) {
            MessageLog ("No disk in drive (" + $Source + "). Copy aborted.`n") ([goVerbosityEnum]::Error)
            return
        }
    
        $VolumeName = $DiskInfo.VolumeName
    }

    if (-Not $UseNewestFolder -And $VolumeName) {
        $DestinationDirectory = $Destination + "\" + $VolumeName
        $DirectoryCreationTime = [DateTime]::Now
    }

    if (-Not $DestinationDirectory) {
        $NewestValidDate=[DateTime]::Now

        if ($SubfolderSearch) {
            get-childitem -Recurse -Directory -Force $Destination | ForEach-Object {
	
	            if ($_.CreationTime -lt $NewestValidDate) { 
		            if ($_.CreationTime -gt $DirectoryCreationTime) {

			            $DestinationDirectory = $_.FullName
			            $DirectoryCreationTime = $_.CreationTime

			            if ($Verbose) {
				             MessageLog ("`tNewest: " + $DirectoryCreationTime + " -- " + $_.Name) ([goVerbosityEnum]::Debug)
			            }
		            }
	            }
            }
        }
        else {
            get-childitem -Directory $Destination | ForEach-Object {
	
	            if ($_.CreationTime -lt $NewestValidDate) { 
		            if ($_.CreationTime -gt $DirectoryCreationTime) {

			            $DestinationDirectory = $_.FullName
			            $DirectoryCreationTime = $_.CreationTime

			            if ($Verbose) {
				             MessageLog ("`tNewest: " + $DirectoryCreationTime + " -- " + $_.Name) ([goVerbosityEnum]::Debug)
			            }
		            }
	            }
            }
        }
    }

    if ($DestinationDirectory -eq "") {
        MessageLog ("No directories found to copy to in " + $Destination + ".`n") ([goVerbosityEnum]::Error)
        exit
    }

    if (-Not $CopyToPopulatedDirectory -and (Test-Path $DestinationDirectory) -and (Get-ChildItem $DestinationDirectory | Measure-Object).Count -gt 0) {
	    MessageLog ("Destination directory (" + $DestinationDirectory + ") is not empty. Copy aborted. Use -CopyToPopulatedDirectory to force write.`n") ([goVerbosityEnum]::Error)
        exit
    }

    $ObjectsToCopy = Get-ChildItem -recurse $Source | Measure-Object -Sum Length

    if ($ObjectsToCopy -eq 0) {
	    MessageLog ("Source directory (" + $Source + ") is empty. Copy aborted.`n") ([goVerbosityEnum]::Error)
        exit
    }
   
    $BytesToCopy = Format-Bytes $ObjectsToCopy.Sum

    $CopyMessage = "Copying " + $ObjectsToCopy.Count + " files, " + $BytesToCopy + ", from " + $Source + " to " + $DestinationDirectory + " (" + $DirectoryCreationTime  + ")`n"

    if ($WhatIfPreference) {
        $CopyMessage = "What if: " + $CopyMessage
    }
    
    MessageLog ($CopyMessage)

    if ($ConfirmPreference -eq "High") {
        $Confirmation = 0
        $SecondsToPause = 10
        ForEach ($Count in (1..$SecondsToPause)) {   
            $Message = "Starting in " + $($SecondsToPause - $Count) + ".."
            Write-Progress -Id 1 -Activity $Message -Status "Ctrl-C to cancel" -PercentComplete (($Count / $SecondsToPause) * 100)
            Start-Sleep -Seconds 1
        }

        Write-Progress -Id 1 -Activity $Message -PercentComplete 100 -Completed
    }
    else {
        $Confirmation = $Host.UI.PromptForChoice("Confirm Copy", "Do you want to copy the "+ $Source + " to " + $DestinationDirectory + "?", ('&Yes', '&No'), 1)
    }
    
    if (-Not $WhatIfPreference -and $Confirmation -eq 0) {
        if (-Not (Test-Path $DestinationDirectory)) {                    
            New-Item -Path $DestinationDirectory -ItemType Directory | out-null
        }	

        $Shell = New-Object -Com Shell.Application
        $Shell.Namespace($DestinationDirectory).CopyHere($Shell.NameSpace($Source).Items())

        MessageLog ("`nCopy complete")
    }
    else {
        MessageLog ("`nCopy skipped")
    }    

    start-sleep (1); 

    if (-Not $NoEject) {
        (New-Object -com "WMPlayer.OCX.7").cdromcollection.item(0).eject()
    }

    if (-Not $Quiet) {
        PlayCompletionSound
    }
}

Function MessageLog($Message, $MessageVerbosity=[goVerbosityEnum]::Information) {
if ($MessageVerbosity.value__ -le $Verbosity.value__) {
         switch ($MessageVerbosity) {
            "Critical"    { Write-Warning $Message; break }
            "Error"       { Write-Warning $Message; break  }
            "Warning"     { Write-Warning $Message; break }
            "Information" { Write-Host $Message; break }
            "Verbose"     { Write-Host $Message; break }
            "Debug"       { Write-Debug $Message;  }
            default       { Write-Host $Message }
        }    
    }
}

Function Format-Bytes {
    Param
    (    
        [Parameter(ValueFromPipeline=$true)]    
        [ValidateNotNullOrEmpty()][float]$number
    )
    Begin{
        $sizes = 'KB','MB','GB','TB','PB'
    }
    Process {
        for($x = 0;$x -lt $sizes.count; $x++){
            if ($number -lt "1$($sizes[$x])"){
                if ($x -eq 0){
                    return "$number B"
                } else {
                    $num = $number / "1$($sizes[$x-1])"
                    $num = "{0:N2}" -f $num                    
                    return "$num $($sizes[$x-1])"
                }
            }
        }
    }

    End{}
}

Function PlayCompletionSound {
[console]::beep(659,500)
[console]::beep(659,500)
[console]::beep(659,500)
[console]::beep(698,350)
[console]::beep(523,150)
[console]::beep(415,500)
[console]::beep(349,350)
[console]::beep(523,150)
[console]::beep(440,1000)
}

Function CheckForUpdate {

    if ($NoAutoUpdate) { return }

    MessageLog ("Version: " + $Version)
    MessageLog ("Checking for updates...")

    $LatestVersion = (New-Object System.Net.WebClient).Downloadstring('https://raw.githubusercontent.com/fpdavis/CopyRemovableMedia/master/Version.txt')
    $LatestVersion = $LatestVersion.Trim().Split(":")
    if ($Version -lt $LatestVersion[0]) {
        MessageLog ("Updating from version " + $Version + " to version " + $LatestVersion[0])

        $LatestVersionPath = $MyInvocation.ScriptName.Replace(".ps1", "_" + $LatestVersion[0] + ".ps1")
        $Script = (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/fpdavis/CopyRemovableMedia/master/CopyRemovableMedia.ps1', $LatestVersionPath)

        $Hash = CalculateHash($LatestVersionPath)

        MessageLog ("Verifying Hash (" + $Hash.Hash + ")") [goVerbosityEnum]::Verbose
        if ($Hash.Hash -eq $LatestVersion[1].ToUpper()) {
            MessageLog ("Hash match") [goVerbosityEnum]::Verbose
            $OldVersionPath = $MyInvocation.ScriptName.Replace(".ps1", "_" + $Version + ".ps1")

            MessageLog ("Archiving current version as " + $OldVersionPath)
            Rename-Item -Path $MyInvocation.ScriptName -NewName $OldVersionPath
            Rename-Item -Path $LatestVersionPath -NewName $MyInvocation.ScriptName

            MessageLog ("Restarting new version...") [goVerbosityEnum]::Warning

            $Paramaters = " -NoAutoUpdate "
            foreach ($Argument in $PSBoundParameters.GetEnumerator()) {
                $Paramaters += " -" + $Argument.Key

                if ($Argument.Value -ne $true) {
                    $Paramaters += " '" + $Argument.Value + "'"
                }
            }
            $MyInvocation.ScriptName + $Paramaters
            Invoke-Expression ($MyInvocation.ScriptName + $Paramaters)
            exit
        }
        else {
           MessageLog ("Hash did not match $($LatestVersion[1]), skipping update.") [goVerbosityEnum]::Error
           Remove-Item $LatestVersionPath
        }
    }
    else {
       MessageLog ("Currently on latest version.")
    }
}

Function CalculateHash($Path) {

    Get-FileHash -Path $Path -Algorithm MD5         
}


Main
