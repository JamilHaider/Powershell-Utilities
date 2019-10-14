<#
.SYNOPSIS
  Set of Utilites for running queries and actions against fleet of devices using sccm  
.DESCRIPTION
  Functions that can be used to validate newly imaged devices with Windows 7 and Windows 10
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


# helper functions
function Get-Out(){
  $Out = [PSCustomObject]@{
    Error = $null
    Result = $null
  }
  $Out
}

# abstract out try catch blocks
function TryRun($Action, $message, $arguments){
  $result = Get-Out
  try {
    $result.Result = $scriptBlock.Invoke($arguments) 
  }
  catch {
    $result.Error = $message
  }
  $result
}

# Retrieve all the packages available in software center for a device
function GetPackages($system){
  $scriptBlock = {
    Get-WmiObject -ComputerName $system -Class CCM_Program -Namespace root\ccm\clientsdk
  }

  $ErrorMessage = "Unable to Read Packages for $system"

  TryRun $scriptBlock $ErrorMessage $system
}

function GetPackage($system, $packageName){ 
  $packages = GetPackages $system
  if($packages.Result){
    $packages.Result = $packages.Result | Where-Object {$_.Name -eq $packageName}
  }
  $packages
}

# Triggers installation on a client system 
function InstallPackage($system, $packageName){

  $package = GetPackage $system $packageName

  $ErrorMessage = "Unable to install the package. Make sure the system is reachable and WMI service WINMGMT is running"

  if($package.Result){
    $scriptBlock = {
      Invoke-WmiMethod -ComputerName $system -Namespace root\ccm\clientsdk -class CCM_ProgramsManager -Name ExecutePrograms -ArgumentList $package.Result
    }
    $package = TryRun $scriptBlock $ErrorMessage $system
  }

  $package
}

# Checks if application is installed as per SCCM client 
function IsAppInstalledSCCM($system, $packageName){
  $out = GetPackage $system $packageName
  if($out.Result){
    $out.Result = $out.Result.LastRunStatus -eq 'Succeeded'
  } 
  $out
}

# delay the installation of updates that come through software center
# usefull in scenario where installation of important packages is being delayed by updates. 
function CancelUpdates($system){
  $scriptBlock = {
    $requests = gwmi -class sms_maintenancetaskrequests -namespace root\ccm -ComputerName $system
    $toCancel = $requests | ? { $_.ClientID -eq "updatesmgr"}
    if($toCancel -ne $null){
      $toCancel.Delete()
      $toCancel.Count + " Updates Cancelled. Restart CCMEXEC to install remaining packages"
    }    
  }

  $ErrorMessage = "Unable to Cancel Updates. Please confirm if CCMEXEC and WINMGMT services are running on the system"

  TryRun $scriptBlock $ErrorMessage $system
}

# Returns all the apps installed and visible in control panel
function GetInstalledApps($system){
  $scriptBlock = {
    Get-InstalledApplication -ComputerName $system
  }

  $ErrorMessage = "Unable to Check Installed Applications for $system"

  TryRun $scriptBlock $ErrorMessage $system
}

# Search for an installed application
function FindInstalledApp($system, $AppName){
  $apps = GetInstalledApps $system
  if($apps.Result){
    $apps.Result = $apps.Result | Where-Object {$_.Application.Contains($AppName)}
  }
  $apps
}

# Creates a Win32_Process on a given system. Non Interactive except for msg.exe
function CreateWin32Process($system, $program){
  $scriptBlock = {
    Invoke-WmiMethod -Class Win32_Process -ComputerName $system -Name Create -ArgumentList $program
  }

  $ErrorMessage = "Unable to Create Process. Please check if system is accessible and you have approprate permissions"

  TryRun $scriptBlock $ErrorMessage $system
}


# Remove MSI package given a guid 
function RemoveMSI($system, $GUID){
  CreateWin32Process $system "msiexec /x $GUID /qn"
}

# Retrieve the asset tag from bios
function GetAssetTag($system){
  $scriptBlock = {
    Get-WmiObject -ComputerName $system Win32_SystemEnclosure 
  }

  $ErrorMessage = "Unable to run WMI command. Make sure system is accessible and WMI is running on the system"

  TryRun $scriptBlock $ErrorMessage $system
}


function DisableBitlocker($system){
  Invoke-Expression "manage-bde.exe -protectors -disable C: -ComputerName $system"  
}

function EnableBitlocker($system){
  Invoke-Expression "manage-bde.exe -protectors -enable C: -ComputerName $system"  
}

# Schedule tasks for execution with automatic deletion after 5 minutes
function ScheduleTask($system, $target, $TaskName){
  $scriptBlock = {
    $time = (Get-Date).AddMinutes(2).ToString("HH:mm")
    # delete task on remote system after delay of 5 Minutes
    $target2 = "sleep 300 && schtasks.exe /Delete /F /TN $TaskName"
    $ntarget = "cmd /c $target && $target2"
    schtasks /create /F /s $system /sc ONCE /RL HIGHEST /TN $TaskName /ST $time  /TR "$ntarget"  
  }

  $ErrorMessage = "Unable to Schedule the task"
  
  TryRun $scriptBlock $ErrorMessage $system
}

# Restart the computer unless the argument is the computer running the script
function RestartComputer($system) {
  if($system -eq $ENV:ComputerName){
      "Skipping $system" 
  }
  else { 
      Restart-Computer -ComputerName $system -Force 
  }
}

# shutdown computer after releasing ip and flushing dns. Useful for situation where lots of computer are 
# plugged in and removed after imaging and there is chance of running out of ip addresses. 
function ShutdownComputer($system) = { 
  if($system -eq $ENV:ComputerName){
      "Skipping $system" 
  }
  else { 
      CreateWin32Process $system "cmd /s /c shutdown /s && ipconfig /release && ipconfig /flushdns"
  }
}

# force group policy update
function GPUpdate($system){
  CreateWin32Process $system 'gpupdate.exe /force'
}

# Shows a popup on remote system.
function SendMessage($system, $message, $timeout = 7200){
  CreateWin32Process $system "C:\Windows\system32\msg.exe /time:$timeout * $message"
}

function PendingReboot($system){
  (Test-PendingReboot -ComputerName $system).IsRebootPending
}

# purgeType = 0 => The next policy request will be for a full policy instead of the change in policy since the last policy request.
# purgeType = 1 => The existing policy will be purged completely.
function ResetPolicy($system, $purgeType = 0){
  Invoke-WmiMethod -ComputerName $system -Namespace root\ccm -Class sms_client -Name ResetPolicy -ArgumentList $purgeType  
}

# Remove pending maintenance task requests which can stop the package in software center to install 
function RemoveRequests($system){
  $request = gwmi -class sms_maintenancetaskrequests -Namespace root\ccm -ComputerName $system
  $requests.Delete()
  Get-Service ccmexec -ComputerName $system  | Restart-Service 
}