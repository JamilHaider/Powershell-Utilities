<#
.SYNOPSIS
  Set of Unitilites for running queries and actions against fleet of devices using sccm  
.DESCRIPTION
  Set of Function that can be used to validate newly imaged devices with Windows 7 and Windows 10
  For example to check
  1. BitLocker Status
  2. Asset Tags
  3. Trigger action for SCCM related packages
  4. Manual trigger installation of SCCM Package
  5. Delay Windows related updates coming via SCCM
  ....and more 
.INPUTS
.OUTPUTS
.NOTES
  Version:        0.1
  Author:         Jamil Haider
  Creation Date:  10/10/2018
  Purpose/Change: Automate validation of O.S and Package deployment
  
.EXAMPLE
#>

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

# Need this to find installed packages
. $ScriptDir\libs\InstalledPrograms.ps1


# helper function
function Get-Out(){
  $Out = [PSCustomObject]@{
    Error = ""
    Result = $null
  }
  $Out
}

# Retrieve all the packages available in software center for a device
function GetPackages($system){
  $result = Get-Out 
  $out = @()
  try {
    $result.Result = Get-WmiObject -ComputerName $system -Class CCM_Program -Namespace root\ccm\clientsdk
  }
  catch {
    $result.Error = "Unable to Read Packages for $system"
  }
  $result
}

# Returns all the apps installed and visible in control panel
function GetInstalledApps($system){
  $result = Get-Out
  try {
    $result.Result = Get-InstalledApplication -ComputerName $system #| Where-Object {$_.Application.Contains("$AppName")}
  }
  catch {
    $result.Error = "Unable to Check Installed Applications for $system"
  }
  $result
  #$out = @()
  #if($res -eq $null){ $out += "$AppName Not Installed"} else { $out += "$AppName is Installed"}
  #$out 
}

# Search for an install application
function FindInstalledApp($system, $AppName){
  $apps = GetInstalledApps $system
  if($apps.Result){
    $apps.Result = $apps.Result | Where-Object {$_.Application.Contains($AppName)}
  }
  $apps
}