# frxprofmgmt
A powershell GUI to manage FSLogix profiles

Requires Active Directory Powershell cmdlets

It searches the profile share for folders named as SamAccountName

Creates a configfile in the same folder as the binary.

'ProfileRoot     ="\\fileserver\profileshare"
The path to the Profile share

'ProfileSubFolder'="\path\below\<username>\in\profileroot\"
The path below the username where the vhdx files reside for example:
\\fileserver\profileshare\username\#sub\folder\#files.vhdx

'pPrefix'         ="Profile_"
The profile filename prefix

'oPrefix'         ="ODFC_"
The office profile filename prefix

'fileExt'         =".VHDX"
The file extension, VHD or VHDX
