# Copy Removable Media
Powershell script to automatically copy removable media when new media is inserted.

## SYNTAX

	CopyRemovableMedia.ps1

		[-Help]
		[-UseNewestFolder]
		[-SubfolderSearch]
		[-CopyToPopulatedDirectory]
		[-Quiet]
		[-NoEject]
		[-ExitAfterCopy]
		[-NoAutoUpdate]
		[-Confirm]
		[-WhatIf]
		[-Source 'c:\Your\Source\Directory']
		[-Destination 'c:\Your\Destination\Directory']
		[-Verbosity Critical|Error|Warning|Information|Verbose|Debug]

## DESCRIPTION

This program was designed to facilitate the manual process of copying and archiving removable media such as CDROM, DVD, and Blu-ray drives. With no parameters specified the program will identify the computers Optical Disk Drive and if there is a disk in the drive it will retrieve the name of the volume and create a new directory with the same name in the current directory. It will then proceed to copy the contents of the entire disk into the newly created directory. A chime is played when the copy is complete to prompt the user to insert another disk for copy.


## PARAMETERS

 - **Help** - Displays this help screen along with removable/optical drive information
and then exits.
- **UseNewestFolder** - By default the program will create a new folder in the Destination
folder with the Volume Name of the drive being copied from. This parameter bypasses the new folder creation and searches the destination folder for the folder with the newest creation date. This folder is then used as the final destination for the file copy. Subfolders will not be searched by default unless the SubfolderSearch parameter is specified. An existing folder with files wills in it will not be written to unless the CopyToPopulatedDirectory parameter is specified.
- **SubfolderSearch** - Searches the subfolders of the Destination directory when looking for the newest folder. Otherwise the newest folder in the root directory will
be chosen.
- **CopyToPopulatedDirectory** - By default the program will only copy files to an empty folder. By specifying this parameter you are giving the program permission to copy files into an existing directory that already contains files. This parameter should be used sparingly and with caution.
- **Quiet** - Refrains from playing a chime when the copy is complete.
-  **Eject** - Refrains from ejecting the optical drive after a copy is complete.
- **ExitAfterCopy** - If this parameter is not specified the program will do multiple copies one after another. If specified the program will exit after the first copy is made.
- **NoAutoUpdate** - Skips checking for an update online.
- **Confirm** - The SupportsShouldProcess argument adds Confirm and WhatIf parameters to the function. The Confirm parameter prompts the user before it performs a copy (and before it creates the destination directory if necessary).
- **WhatIf** - The SupportsShouldProcess argument adds Confirm and WhatIf parameters to
the function. The WhatIf parameter lists the changes that the command would make, instead of running the commands.
- **Source** - By default the program identifies the optical drive to copy from. This can be overridden. with a different drive and even a complete file path.
- **Destination** - This is the current directory by default. This is the directory where a new folder will be created based on the Volume Name of the Source Disk. If the UseNewestFolder is specified then this is the directory that will be searched for the newest empty folder to copy the source contents into.
- **Verbosity** - Displays various levels of messaging.
 -- Debug - Super detail
 -- Verbose – Everything Important
 -- Information - Updates that might be useful to user
 -- Warning - Bad things that happen, but are expected
 -- Error - Expected errors that are not recoverable
 -- Critical - Unexpected errors that stop execution

## NOTES

- Does not recurse into symbolically linked directories.
- Limited support for USB and floppy drives. These drive letters must be explicitly specified as the Source.


