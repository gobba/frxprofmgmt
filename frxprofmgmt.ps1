function New-Config {
    #Checks for target directory and creates if non-existent 
		
	#Setup default preferences	
	#Creates hash table and .clixml config file
    $Config = @{
        'ProfileRoot'     ="\\path\to\profile\share"
        'ProfileSubFolder'="\path\below\<username>\in\profileroot\"
        'pPrefix'         ="Profile_"
        'oPrefix'         ="ODFC_"
        'fileExt'         =".VHDX"

    }
    $Config | Export-Clixml -Path ".\frxprofmgmt.config"
	Import-Config
} #end function New-Config

function Import-Config {
	#If a config file exists for the current user in the expected location, it is imported
	#and values from the config file are placed into global variables
	if (Test-Path -Path ".\frxprofmgmt.config") {
		try {
			#Imports the config file and saves it to variable $Config
			$global:Config = Import-Clixml -Path ".\frxprofmgmt.config"
			
			#Creates global variables for each config property and sets their values
            
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("An error occurred importing your Config file. A new Config file will be generated for you. $_", 'Import Config Error', 'OK', 'Error')
			New-Config
		}
	} #end if config file exists
	else {
		New-Config
	}
} #end function Import-Config

function Update-Config {
    #Creates a new Config hash table with the current preferences set by the user
	$Config = @{
		'ProfileRoot'     =$profileRoot
        'ProfileSubFolder'=$profileSubFolder
        'pPrefix'         =$pPrefix
        'oPrefix'         =$oPrefix
        'fileExt'         =$fileExt
	}
    #Export the updated config
    $Config | Export-Clixml -Path ".\frxprofmgmt.config"
} #end function Update-Config

Import-Config

$oPrefix = $global:Config.oprefix
$pPrefix = $global:config.pPrefix
$fileExt = $global:config.fileExt
$profileRoot = $global:config.Profileroot
$profileSubFolder = $global:config.ProfileSubFolder


Add-Type -assembly System.Windows.Forms


# Start creating main form
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='FSLogix Profile Management Tool'
$main_form.Width = 600
$main_form.Height = 400
$main_form.AutoSize = $true

$lblCombBox = New-Object System.Windows.Forms.Label
$lblCombBox.Text = "Profiles"
$lblCombBox.Location  = New-Object System.Drawing.Point(0,10)
$lblCombBox.AutoSize = $true

$ComboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

$comboBox.AutoCompleteSource = [system.windows.forms.AutoCompleteSource]::ListItems
$ComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend

$ComboBox.Width = 300
# Find profile folders
$profiles = Get-ChildItem $profileRoot -ErrorAction SilentlyContinue
Foreach ($profile in $profiles)
{
    $uname = $profile.Name
    # Verify the user exists in AD
    if(get-aduser -Filter{ SamAccountName -eq $uname} )
    {
        $ComboBox.Items.Add($profile.Name)
        
    }

}
$ComboBox.Location  = New-Object System.Drawing.Point(60,10)
$comboBox.Sorted = $true




$lblUserText = New-Object System.Windows.Forms.Label
$lblUserText.Text = "User:"
$lblUserText.Location  = New-Object System.Drawing.Point(0,40)
$lblUserText.AutoSize = $true

$lblUserField = New-Object System.Windows.Forms.Label
$lblUserField.Text = ""
$lblUserField.Location  = New-Object System.Drawing.Point(110,40)
$lblUserField.AutoSize = $true

$lblTotSizeText = New-Object System.Windows.Forms.Label
$lblTotSizeText.Text = "Total Size:"
$lblTotSizeText.Location  = New-Object System.Drawing.Point(0,80)
$lblTotSizeText.AutoSize = $true

$lblTotSizeField = New-Object System.Windows.Forms.Label
$lblTotSizeField.Text = ""
$lblTotSizeField.Location  = New-Object System.Drawing.Point(110,80)
$lblTotSizeField.AutoSize = $true

$lblProfSizeText = New-Object System.Windows.Forms.Label
$lblProfSizeText.Text = "Profile Size:"
$lblProfSizeText.Location  = New-Object System.Drawing.Point(0,120)
$lblProfSizeText.AutoSize = $true

$lblProfSizeField = New-Object System.Windows.Forms.Label
$lblProfSizeField.Text = ""
$lblProfSizeField.Location  = New-Object System.Drawing.Point(110,120)
$lblProfSizeField.AutoSize = $true

$lblOfficeSizeText = New-Object System.Windows.Forms.Label
$lblOfficeSizeText.Text = "Office Size:"
$lblOfficeSizeText.Location  = New-Object System.Drawing.Point(0,140)
$lblOfficeSizeText.AutoSize = $true

$lblOfficeSizeField = New-Object System.Windows.Forms.Label
$lblOfficeSizeField.Text = ""
$lblOfficeSizeField.Location  = New-Object System.Drawing.Point(110,140)
$lblOfficeSizeField.AutoSize = $true

$BtnConfigure = New-Object System.Windows.Forms.Button
$BtnConfigure.Location = New-Object System.Drawing.Size(400,10)
$BtnConfigure.Size = New-Object System.Drawing.Size(120,23)
$BtnConfigure.Text = "Configure"

$btnDelAll = New-Object System.Windows.Forms.Button
$btnDelAll.Location = New-Object System.Drawing.Size(400,40)
$btnDelAll.Size = New-Object System.Drawing.Size(120,23)
$btnDelAll.Text = "Delete All"
$btnDelAll.Enabled = $false

$btnDelProf = New-Object System.Windows.Forms.Button
$btnDelProf.Location = New-Object System.Drawing.Size(400,120)
$btnDelProf.Size = New-Object System.Drawing.Size(120,23)
$btnDelProf.Text = "Delete Profile"
$btnDelProf.Enabled = $false

$btnDelOffice = New-Object System.Windows.Forms.Button
$btnDelOffice.Location = New-Object System.Drawing.Size(400,140)
$btnDelOffice.Size = New-Object System.Drawing.Size(120,23)
$btnDelOffice.Text = "Delete Office"
$btnDelOffice.Enabled = $false

$btnConfSave = New-Object System.Windows.Forms.Button
$btnConfSave.Location = New-Object System.Drawing.Size(400,40)
$btnConfSave.Size = New-Object System.Drawing.Size(120,23)
$btnConfSave.Text = "Save"
$btnConfSave.Enabled = $true

# Refresh iterface on combobox changes
$ComboBox_SelectedIndexChanged=
{

    $selProfile = $profiles.where({$_.name -eq $ComboBox.SelectedItem})
    $user = get-aduser $selProfile.name -ErrorAction SilentlyContinue
    $size = "{0} MB" -f ((Get-ChildItem $selprofile.fullname -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
    $profileFile = $selProfile.fullname + "$profilesubfolder$pPrefix" + $selProfile.name + $fileExt
    $officeFile = $selProfile.fullname + "$profilesubfolder$oPrefix" + $selProfile.name + $fileExt
    $lblUserField.Text =  $user.name
    $lblTotSizeField.Text =  $size
    $pSize = (Get-ChildItem $profileFile -ErrorAction SilentlyContinue| Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
    $oSize = (Get-ChildItem $officeFile -ErrorAction SilentlyContinue| Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
    $lblProfSizeField.text = "{0} MB" -f ($pSize)
    $lblOfficeSizeField.text = "{0} MB" -f ($oSize)
    
    if($pSize -or $oSize)
    {

        $btnDelAll.Enabled = $true
    }
    if(Get-ChildItem $profileFile -ErrorAction SilentlyContinue)
    {
        $btnDelProf.Enabled = $true
    }else{
        $btnDelProf.Enabled = $false
    }
    if(Get-ChildItem $officeFile -ErrorAction SilentlyContinue)
    {
        $btnDelOffice.Enabled = $true
    }else{
        $btnDelOffice.Enabled = $false
    }

}

$main_form.Controls.Add($ComboBox)
$ComboBox.add_SelectedIndexChanged($Combobox_selectedindexchanged)


$btnDelProf.add_click(
{
    $selProfile = $profiles.where({$_.name -eq $ComboBox.SelectedItem})
    $profileFile = $selProfile.fullname + "$profilesubfolder$pPrefix" + $selProfile.name + $fileExt
    $datenow= [string](Get-Date -Format ddMMyyHHmm)
    $newfilename = "$profileFile-$datenow"
    $lblProfSizeField.Text = "Deleted"
    Rename-Item $profileFile $newfilename
}
)

$btnDelOffice.add_click(
{
    $selProfile = $profiles.where({$_.name -eq $ComboBox.SelectedItem})
    $officeFile = $selProfile.fullname + "$profilesubfolder$oPrefix" + $selProfile.name + $fileExt
    $datenow= [string](Get-Date -Format ddMMyyHHmm)
    $newfilename = "$officeFile-$datenow"
    $lblOfficeSizeField.Text = "Deleted"
    Rename-Item $officeFile $newfilename
}
)

$btnDelAll.Add_Click(
{
    

    $selProfile = $profiles.where({$_.name -eq $ComboBox.SelectedItem})
    #$path = $selProfile.fullname
    $datenow= [string](Get-Date -Format ddMMyyHHmm)
    $officeFile = $selProfile.fullname + "$profilesubfolder$oPrefix" + $selProfile.name + $fileExt
    $newfilename = "$officeFile-$datenow"
    $lblOfficeSizeField.Text = "Deleted"
    Rename-Item $officeFile $newfilename

    $profileFile = $selProfile.fullname + "$profilesubfolder$pPrefix" + $selProfile.name + $fileExt
    $newfilename = "$profileFile-$datenow"
    $lblProfSizeField.Text = "Deleted"
    Rename-Item $profileFile $newfilename
        
    $btnDelAll.Enabled = $false
    $btnDelProf.Enabled = $false
    $btnDelOffice.Enabled = $false

    
    $combobox.Items.Clear()
    $profiles = Get-ChildItem $profileRoot -ErrorAction SilentlyContinue

    Foreach ($profile in $profiles)
    {
        $uname = $profile.Name
        if(get-aduser -Filter{ SamAccountName -eq $uname} )
        {
            $ComboBox.Items.Add($profile.Name)
        
        }

    }
}

)

$BtnConfigure.add_click(
{
    $config_form.Visible=$true
}
)



$config_form                     = New-Object system.Windows.Forms.Form
$config_form.ClientSize          = '600,400'
$config_form.text                = "FSLogix Profile Management Tool"
$config_form.TopMost             = $false
$config_form.Visible             = $false
$config_form.ControlBox          = $false

$txtbConfProfileRoot             = New-Object system.Windows.Forms.TextBox
$txtbConfProfileRoot.multiline   = $false
$txtbConfProfileRoot.text        = "$profileroot"
$txtbConfProfileRoot.width       = 300
$txtbConfProfileRoot.height      = 20
$txtbConfProfileRoot.location    = New-Object System.Drawing.Point(134,16)
$txtbConfProfileRoot.Font        = 'Microsoft Sans Serif,10'

$lblConfProfileRoot              = New-Object system.Windows.Forms.Label
$lblConfProfileRoot.text         = "Profile Root"
$lblConfProfileRoot.AutoSize     = $true
$lblConfProfileRoot.width        = 25
$lblConfProfileRoot.height       = 10
$lblConfProfileRoot.location     = New-Object System.Drawing.Point(14,20)
$lblConfProfileRoot.Font         = 'Microsoft Sans Serif,10'

$btnConfSave                     = New-Object system.Windows.Forms.Button
$btnConfSave.text                = "Save"
$btnConfSave.width               = 60
$btnConfSave.height              = 30
$btnConfSave.location            = New-Object System.Drawing.Point(446,322)
$btnConfSave.Font                = 'Microsoft Sans Serif,10'
$lblConfPrfSubFolder             = New-Object system.Windows.Forms.Label
$lblConfPrfSubFolder.text        = "Sub Folder"
$lblConfPrfSubFolder.AutoSize    = $true
$lblConfPrfSubFolder.width       = 25
$lblConfPrfSubFolder.height      = 10
$lblConfPrfSubFolder.location    = New-Object System.Drawing.Point(14,51)
$lblConfPrfSubFolder.Font        = 'Microsoft Sans Serif,10'

$txtbConfSubFolder               = New-Object system.Windows.Forms.TextBox
$txtbConfSubFolder.multiline     = $false
$txtbConfSubFolder.text          = "$profileSubFolder"
$txtbConfSubFolder.width         = 100
$txtbConfSubFolder.height        = 20
$txtbConfSubFolder.location      = New-Object System.Drawing.Point(134,47)
$txtbConfSubFolder.Font          = 'Microsoft Sans Serif,10'

$lblConfpPrefix                  = New-Object system.Windows.Forms.Label
$lblConfpPrefix.text             = "Profile Prefix"
$lblConfpPrefix.AutoSize         = $true
$lblConfpPrefix.width            = 25
$lblConfpPrefix.height           = 10
$lblConfpPrefix.location         = New-Object System.Drawing.Point(14,80)
$lblConfpPrefix.Font             = 'Microsoft Sans Serif,10'

$txtbConfpPrefix                 = New-Object system.Windows.Forms.TextBox
$txtbConfpPrefix.multiline       = $false
$txtbConfpPrefix.text            = "$pPrefix"
$txtbConfpPrefix.width           = 100
$txtbConfpPrefix.height          = 20
$txtbConfpPrefix.location        = New-Object System.Drawing.Point(134,76)
$txtbConfpPrefix.Font            = 'Microsoft Sans Serif,10'

$txtbConfoPrefix                 = New-Object system.Windows.Forms.TextBox
$txtbConfoPrefix.multiline       = $false
$txtbConfoPrefix.text            = "$oPrefix"
$txtbConfoPrefix.width           = 100
$txtbConfoPrefix.height          = 20
$txtbConfoPrefix.location        = New-Object System.Drawing.Point(134,107)
$txtbConfoPrefix.Font            = 'Microsoft Sans Serif,10'

$lblConfoPrefix                  = New-Object system.Windows.Forms.Label
$lblConfoPrefix.text             = "Office Prefix"
$lblConfoPrefix.AutoSize         = $true
$lblConfoPrefix.width            = 25
$lblConfoPrefix.height           = 10
$lblConfoPrefix.location         = New-Object System.Drawing.Point(14,111)
$lblConfoPrefix.Font             = 'Microsoft Sans Serif,10'

$lblConfFileExt                  = New-Object system.Windows.Forms.Label
$lblConfFileExt.text             = "File Extension"
$lblConfFileExt.AutoSize         = $true
$lblConfFileExt.width            = 25
$lblConfFileExt.height           = 10
$lblConfFileExt.location         = New-Object System.Drawing.Point(14,140)
$lblConfFileExt.Font             = 'Microsoft Sans Serif,10'

$txtbConfFileExt                 = New-Object system.Windows.Forms.TextBox
$txtbConfFileExt.multiline       = $false
$txtbConfFileExt.text            = "$fileExt"
$txtbConfFileExt.width           = 100
$txtbConfFileExt.height          = 20
$txtbConfFileExt.location        = New-Object System.Drawing.Point(134,136)
$txtbConfFileExt.Font            = 'Microsoft Sans Serif,10'

$config_form.controls.AddRange(@($txtbConfProfileRoot,$lblConfProfileRoot,$btnConfSave,$lblConfPrfSubFolder,$txtbConfSubFolder,$lblConfpPrefix,$txtbConfpPrefix,$txtbConfoPrefix,$lblConfoPrefix,$lblConfFileExt,$txtbConfFileExt))


$btnConfSave.Add_Click(
{ 
    $profileRoot = $txtbConfProfileRoot.text
    Update-Config
    $config_form.Visible=$false 
        $combobox.Items.Clear()
    $profiles = Get-ChildItem $profileRoot -ErrorAction SilentlyContinue

    Foreach ($profile in $profiles)
    {
        $uname = $profile.Name
        if(get-aduser -Filter{ SamAccountName -eq $uname} )
        {
            $ComboBox.Items.Add($profile.Name)
        
        }

    }


    

}
)








$main_form.Controls.Add($BtnConfigure)
$main_form.Controls.Add($btnDelAll)
$main_form.Controls.Add($btnDelProf)
$main_form.Controls.Add($btnDelOffice)
$main_form.Controls.Add($lblUserText)
$main_form.Controls.Add($lblUserField)
$main_form.Controls.Add($lblTotSizeText)
$main_form.Controls.Add($lblTotSizeField)
$main_form.Controls.Add($lblProfSizeText)
$main_form.Controls.Add($lblProfSizeField)
$main_form.Controls.Add($lblOfficeSizeText)
$main_form.Controls.Add($lblOfficeSizeField)

$main_form.Controls.Add($lblCombBox)
$main_form.BringToFront()
$main_form.ShowDialog()



