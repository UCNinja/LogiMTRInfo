


# Script that gathers info for Logitech Microsoft Teams Room Systems
# GetSoftware function from many sources including
# https://gallery.technet.microsoft.com/scriptcenter/Get-Software-Function-to-bd2e0204

# Author: Luke Kannel
# Version Control:
# V1.0 - Initial Release
# V1.1 - Updated Skype Room Enumeration to look for any version
# V1.2 - Updated to script-scoped variables for future data parsing and export
# V2.0 - Removed explicit Surface Pro check, now reports all small form factors. Added driver versions, services, and Windows Hotfixes

Param(
    [parameter(Mandatory=$false,
    ParameterSetName="Computer")]
    [String[]]
    $ComputerName

)

function InitVariables() {
  $script:HTMLContent = $null
  $HardwareHTML = $null
  $BIOSHTML = $null
	$script:HardwareInfo  = @()
  $script:BIOSInfo  = @()
	$script:WindowsInfo  = @()
	$script:SkypeRoomInfo = @()
	$script:LogiSoftware = @()
	$script:DriverInfo = @()
	$script:ConnectedUSBDevicesInfo = @()	
	$script:ServicesInfo = @()
	$script:HotfixInfo = @()
}



function GetComputerBaseline() {

  $SRSInfo = Get-ComputerInfo
  return $SRSInfo
}


function GetHardwareInfo() {

  $script:HardwareInfo = gwmi -Class Win32_ComputerSystem
  $script:BIOSInfo = gwmi -Class Win32_Bios
  $script:DriverInfo = Get-WmiObject Win32_PnPSignedDriver | Select Description,DeviceName,DriverVersion,DriverDate

}

function GetWindowsVersion() {
$WinVer = New-Object -TypeName PSObject
$WinVer | Add-Member -MemberType NoteProperty -Name Major -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentMajorVersionNumber).CurrentMajorVersionNumber
$WinVer | Add-Member -MemberType NoteProperty -Name Minor -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentMinorVersionNumber).CurrentMinorVersionNumber
$WinVer | Add-Member -MemberType NoteProperty -Name Build -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' CurrentBuild).CurrentBuild
$WinVer | Add-Member -MemberType NoteProperty -Name Revision -Value $(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion' UBR).UBR
$SRSFullWindowsVersion = $WinVer.Major, $WinVer.Minor, $WinVer.Build, $WinVer.Revision -join "."

$script:WindowsInfo = New-Object -TypeName PSObject
$script:WindowsInfo | Add-Member -MemberType NoteProperty -Name WindowsVersion -Value $SRSFullWindowsVersion


# Services Info. Need to create separate software function for Windows Version, Windows Info, Services, Hotfixes and other Software (Like WU Info)

$script:ServicesInfo = Get-Service | Sort DisplayName

$script:ConnectedUSBDevicesInfo = Get-PnpDevice -PresentOnly | Where-Object { $_.InstanceId -match '^USB' } | sort -Property FriendlyName -Descending

$script:HotfixInfo = get-wmiobject -class win32_quickfixengineering

}




function CheckWindowsActivation() {
$SRSLicenseObject = Get-CimInstance -ClassName SoftwareLicensingProduct |where PartialProductKey |select LicenseStatus
If ($SRSLicenseObject.LicenseStatus -ne "1") {
	$SRSLicenseStatus = "Not Activated" } else {
	$SRSLicenseStatus = "Activated"
	}

$script:WindowsInfo | Add-Member -MemberType NoteProperty -Name ActivationStatus -Value $SRSLicenseStatus
#Write-Host "Windows Activation Status:" $SRSLicenseStatus
}

function GetSRSVersion() {

$script:SkypeRoomInfo = New-Object -TypeName PSObject

#Modified to use query from Technet for determining SRS existence.

$package = get-appxpackage -Name Microsoft.SkypeRoomSystem -User Skype; if ($package -eq $null) {
  $script:SkypeRoomInfo | Add-Member -MemberType NoteProperty -Name SkypeRoomVersion -Value "NotInstalled"

	} else {
	$SRSVersion = $package.Version
  $script:SkypeRoomInfo | Add-Member -MemberType NoteProperty -Name SkypeRoomVersion -Value $SRSVersion


	}

#$SRSInfo | Add-Member -NotePropertyName "SRSVersion" -NotePropertyValue  $SRSVersion

}

Function GetSoftware  {
  [OutputType('System.Software.Inventory')]
  [Cmdletbinding()]
  Param(
  [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
  [String[]]$Computername=$env:COMPUTERNAME
  )

  Begin {

  }

  Process  {

  #Init variables
  $temp = $null

  	ForEach  ($Computer in  $Computername){
  		If  (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {
  			$Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
  				ForEach($Path in $Paths) {
					Write-Verbose  "Checking Path: $Path"

          #Init variables
          $temp = $null

  #  Create an instance of the Registry Object and open the HKLM base key

  					Try  {
  					$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64')
						} Catch  {
						Write-Error $_
						Continue
						}

  #  Drill down into the Uninstall key using the OpenSubKey Method

  					Try  {
						$regkey=$reg.OpenSubKey($Path)
	# Retrieve an array of string that contain all the subkey names

						$subkeys=$regkey.GetSubKeyNames()

  # create Array for storing values


  # Open each Subkey and use GetValue Method to return the required  values for each


  					ForEach ($key in $subkeys){
  					Write-Verbose "Key: $Key"
  					$thisKey=$Path+"\\"+$key

						Try {
						$thisSubKey=$reg.OpenSubKey($thisKey)

  # Prevent Objects with empty DisplayName

						$DisplayName =  $thisSubKey.getValue("DisplayName")
  					If ($DisplayName  -AND $DisplayName -like "Logitech*" -AND $DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {
						$Date = $thisSubKey.GetValue('InstallDate')
						If ($Date) {
							Try {
							$Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
							} Catch{
							Write-Warning "$($Computer): $_ <$($Date)>"
  					$Date = $Null
							}
						}

  # Create New Object with empty Properties

  $Publisher =  Try {
  	$thisSubKey.GetValue('Publisher').Trim()
  	}
  	Catch {
  	$thisSubKey.GetValue('Publisher')
  	}
  $Version = Try {
  	#Some weirdness with trailing [char]0 on some strings
  	$thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
  	}
  	Catch {
  	$thisSubKey.GetValue('DisplayVersion')
  	}

$UninstallString =  Try {
  $thisSubKey.GetValue('UninstallString').Trim()
  }
  Catch {
  $thisSubKey.GetValue('UninstallString')
  }

  $InstallLocation =  Try {
  $thisSubKey.GetValue('InstallLocation').Trim()
  }
  Catch {
  $thisSubKey.GetValue('InstallLocation')
  }

  $InstallSource =  Try {
  	$thisSubKey.GetValue('InstallSource').Trim()
  	}
  Catch {
  	$thisSubKey.GetValue('InstallSource')
  	}

  $HelpLink = Try {
  $thisSubKey.GetValue('HelpLink').Trim()
  }
  Catch {
  $thisSubKey.GetValue('HelpLink')
  }

  $Object = [pscustomobject]@{
		Computername = $Computer
  	DisplayName = $DisplayName
  	Version  = $Version
  	InstallDate = $Date
  	Publisher = $Publisher
  	UninstallString = $UninstallString
  	InstallLocation = $InstallLocation
  	InstallSource  = $InstallSource
  	HelpLink = $thisSubKey.GetValue('HelpLink')
  	EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
  }

  $Object.pstypenames.insert(0,'System.Software.Inventory')


$temp = New-Object System.Object
$temp |  Add-Member -NotePropertyName Software -NotePropertyValue $DisplayName -PassThru | Add-Member -NotePropertyName SoftwareVersion -NotePropertyValue $Version -PassThru
$script:LogiSoftware += $temp


#$SRSInfo | Add-Member -NotePropertyName "LogiCamVersion" -NotePropertyValue  $Version

  }

  } Catch {
  Write-Warning "$Key : $_"
  }
  }
  } Catch  {}
  $reg.Close()

  }
  } Else  {
  Write-Error  "$($Computer): unable to reach remote system!"
  }
  	}
  }
}

function OutputContent() {

#Initialize Function variables

$HardwareHTML = $null
$WindowsHTML = $null
$SkypeRoomHTML = $null
$LogiSoftwareHTML = $null
$DriverInfoHTML = $null
$ConnectedUSBDevicesInfoHTML = $null
$ServicesInfoHTML = $null
$HotfixInfoHTML = $null

#Format HTML Header (thanks https://4sysops.com/archives/building-html-reports-in-powershell-with-convertto-html/)
# And thanks to https://techontip.wordpress.com/2015/01/08/powershell-html-report-with-multiple-tables/

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$HardwareHTML = ($script:HardwareInfo | ConvertTo-HTML -As LIST -Property Name,PartofDomain,Domain,Workgroup,Manufacturer,Model -Fragment -PreContent '<h2>Computer Info</h2>' | Out-String )
$BIOSHTML = ($script:BIOSInfo | ConvertTo-HTML -As LIST -Property SerialNumber,SMBIOSBIOSVersion -Fragment -PreContent '<h2>BIOS Info</h2>' | Out-String )
$WindowsHTML = ($script:WindowsInfo | ConvertTo-HTML -As LIST -Property WindowsVersion,ActivationStatus -Fragment -PreContent '<h2>Windows Info</h2>' | Out-String )
$SkypeRoomHTML = ($script:SkypeRoomInfo | ConvertTo-HTML -As LIST -Property SkypeRoomVersion -Fragment -PreContent '<h2>Skype Room Info</h2>' | Out-String )
$LogiSoftwareHTML = ($script:LogiSoftware | ConvertTo-HTML -Property Software,SoftwareVersion -Fragment -PreContent '<h2>Logitech Software Info</h2>' | Out-String )
$DriverInfoHTML = ($script:DriverInfo | ConvertTo-HTML -Property Description,DeviceName,DriverVersion,DriverDate -Fragment -PreContent '<h2>Driver Info</h2>' | Out-String )
$ConnectedUSBDevicesInfoHTML = ($script:ConnectedUSBDevicesInfo | ConvertTo-HTML -Property FriendlyName,Status,Class,InstanceID -Fragment -PreContent '<h2>Connected USB Devices Info</h2>' | Out-String )
$ServicesInfoHTML = ($script:ServicesInfo | ConvertTo-HTML -Property DisplayName,Status -Fragment -PreContent '<h2>Services Info</h2>' | Out-String )
$HotfixInfoHTML = ($script:HotfixInfo | ConvertTo-HTML -Property Description,HotFixID,InstalledOn -Fragment -PreContent '<h2>Hotfix Info</h2>' | Out-String )


#Setup Output File
if (Test-Path C:\Logitech) {} Else {mkdir C:\Logitech}
$FileName = $null
$OutputFileName = $null
$FileName = (Get-Date).tostring("dd-MM-yyyy-hh-mm-ss")
$OutputFileName = ("C:\Logitech\" + ($script:HardwareInfo.Name) + "_" + ($FileName) + ".html")

ConvertTo-HTML -Head $Header -PostContent $HardwareHTML,$BIOSHTML,$WindowsHTML,$SkypeRoomHTML,$LogiSoftwareHTML,$DriverInfoHTML,$ConnectedUSBDevicesInfoHTML,$ServicesInfoHTML,$HotfixInfoHTML -PreContent "<h1>Logitech Inventory</h1>" | Out-File $OutputFileName

Start $OutputFileName

}

function WrapItUp() {

#Read-Host "Press any key to exit"

}


function Main() {

If ($ComputerName) {
	Write-Host $ComputerName
	$Credentials = Get-Credential
	#Run Remote Powershell

	Invoke-Command -ComputerName $ComputerName -ScriptBlock ${Function:GetHardwareInfo} -Credential $Credentials
	GetWindowsVersion
	CheckWindowsActivation
	GetSRSVersion
	GetSoftware
	OutputContent
	WrapItUp

	} else {
	Write-Host "Running against local computer..."
	#Run local powershell
	GetHardwareInfo
	GetWindowsVersion
	CheckWindowsActivation
	GetSRSVersion
	GetSoftware
	OutputContent
	WrapItUp
	}

}

InitVariables
Main
