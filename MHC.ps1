### Check if the folder Results exists. If Exists then it will
### delete it and create new empty folders
####################################################################

$destination = ".\Results"
If (Test-Path $destination){
    Remove-Item -path ".\Results" -recurse
}
New-Item -Path ".\Results" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item -Path ".\Results\MHC_Excel" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
New-Item -Path ".\Results\MHC_Info" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null


# Bypass  policy #############
$Bypass = Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
$Bypass
####################################################################

## Initial values for checking Before, After and Sass scan ##########

####################################################################

#######   EXCEL Info  ##############################################

## Instantiate the COM object

$excel_file_path = 'C:\Scripts\MHC\Template_MHC.xls'
$Excel = New-Object -ComObject Excel.Application
$date = Get-Date -DisplayHint Date
#######  Initial values according to Technical Specification

$Init_PSW_Length = 8
$Init_PSW_CMPL = 'min=disabled,8,8,8,8 passphrase=0'
$Init_Lock_Duration = '60'
$Init_fails = '10'
$Init_retention = '90000'
$Init_Network = 'False'
$Init_Stopped = 'Stopped'
$Init_Running = 'Running'
$Init_policy = 'Start and Stop with Host'
$Init_timmer = '600'

##  Create menu  ############
 cls
     
   $vcenter = $Server
  # $esxi = $S_esxi 
  $credential
  Read-Host
   Connect-VIServer -Server $Server -Credential $credential
   

foreach($line in Get-Content .\list.txt) 
{
 if($line -match $regex){
   $esxi =$line

 ####################################################################

 ##  Check  the Security Policy ##########
 
 ####################################################################

 # Check number of failed attemps before lock account.
   $acclock = Get-AdvancedSetting -Name "Security.AccountLockFailures"$esxi  | select Value -ExpandProperty Value
 ####################################################################

  # Check minutes of lockout duration.
   $timelock = Get-AdvancedSetting -Name "Security.AccountUnlockTime"$esxi  | select Value -ExpandProperty Value
 ####################################################################
 
 # Check password length.
  ####################################################################

  # Check password complexity
   $complexity = Get-AdvancedSetting -Name "Security.PasswordQualityControl"$esxi  | select Value -ExpandProperty Value
  ####################################################################
  
 # Log retaition 
   $rotate = Get-AdvancedSetting -Name "Syslog.global.defaultRotate" $esxi  | select Value -ExpandProperty Value 
   $size = Get-AdvancedSetting -Name "Syslog.global.defaultSize" $esxi  | select Value -ExpandProperty Value
   $total = $rotate * $size
  ####################################################################
   
 # Remote sys log 
   $syslogglobal = Get-AdvancedSetting -Name "Syslog.global.logHost" $esxi  | select Value -ExpandProperty Value
     $syslogglobal_set = $syslogglobal
   if ($syslogglobal -eq "") {
     $syslogglobal_set = "Global syslog is not set"
   }
   ####################################################################

  # Check open ports
  $esxiport = Get-VMHostFirewallException -VMHost $esxi -Enabled:$True | Format-Table
  ####################################################################

# Check SMTP service
   $smtp_status =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq "snmpd"} | select Running -ExpandProperty Running
   $smtp_policy =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'snmpd'} | select policy -ExpandProperty Policy 
   
   ## Check the status and set value to stopped or Running
   if ($smtp_status) {
     $smtp_st = "Running"  
   }   
   else {
     $smtp_st = "Stopped"
   }

   ## Check the policy 
   if ($smtp_policy -contains "on") {
     $smtp_po = "Start and Stop with Host"  
   }   
   elseif ($tsmssh_policy -contains "off") {
     $smtp_po = "Start and Stop Manually"
   }
   else {
     $smtp_po = "Start and Stop with port Usage"
   }

####################################################################

# Check Remote Tech Support (SSH) Service
   $tsm_status =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'TSM'} | select Running -ExpandProperty Running
   $tsm_policy =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'TSM'} | select policy -ExpandProperty Policy 
   
   ## Check the status and set value to stopped or Running
   if ($tsm_status) {
     $tsm_st = "Running"  
   }   
   else {
     $tsm_st = "Stopped"
   }

   ## Check the policy 
   if ($tsm_policy -contains "on") {
     $tsm_po = "Start and Stop with Host"  
   }   
   elseif ($tsm_policy -contains "off") {
     $tsm_po = "Start and Stop Manually"
   }
   else {
     $tsm_po = "Start and Stop with port Usage"
   }
    
####################################################################

# Check Local Tech Support Service
   $tsmssh_status =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'TSM-SSH'} | select Running -ExpandProperty Running
   $tsmssh_policy =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'TSM-SSH'} | select policy -ExpandProperty Policy 
   
   ## Check the status and set value to stopped or Running
   if ($tsmssh_status) {
     $tsmssh_st = "Running"  
   }   
   else {
     $tsmssh_st = "Stopped"
   }

   ## Check the policy 
   if ($tsmssh_policy -contains "on") {
     $tsmssh_po = "Start and Stop with Host"  
   }   
   elseif ($tsmssh_policy -contains "off") {
     $tsmssh_po = "Start and Stop Manually"
   }
   else {
     $tsmssh_po = "Start and Stop with port Usage"
   }
     
####################################################################

# Check NTP Daemon
   $ntp_status =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'ntpd'} | select Running -ExpandProperty Running
   $ntp_policy =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'ntpd'} | select policy -ExpandProperty Policy 
   
   ## Check the status and set value to stopped or Running
   if ($ntp_status) {
     $ntp_st = "Running"  
   }   
   else {
     $ntp_st = "Stopped"
   }

   ## Check the policy 
   if ($ntp_policy -contains "on") {
     $ntp_po = "Start and Stop with Host"  
   }   
   elseif ($ntp_policy -contains "off") {
     $ntp_po = "Start and Stop Manually"
   }
   else {
     $ntp_po = "Start and Stop with port Usage"
   }
     
####################################################################

# Check DCUI 
   $dcui_status =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'DCUI'} | select Running -ExpandProperty Running
   $dcui_policy =   Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq 'DCUI'} | select policy -ExpandProperty Policy 
   
   ## Check the status and set value to stopped or Running
   if ($dcui_status) {
     $dcui_st = "Running"  
   }   
   else {
     $dcui_st = "Stopped"
   }

   ## Check the policy 
   if ($dcui_policy -contains "on") {
     $dcui_po = "Start and Stop with Host"  
   }   
   elseif ($dcui_policy -contains "off") {
     $dcui_po = "Start and Stop Manually"
   }
   else {
     $dcui_po = "Start and Stop with port Usage"
   }
     
 ####################################################################

 # DCUI timeout value.
   $dcui_timer = Get-AdvancedSetting -Name "UserVars.DcuiTimeOut" $esxi  | select Value -ExpandProperty Value
 ####################################################################
 
 # Lockdown mode.
 ####################################################################
   $lockdown = Get-VMHost $esxi | Select Name,@{N="Lockdown";E={$_.Extensiondata.Config.adminDisabled}}| Select Lockdown -ExpandProperty Lockdown
      
   ## Check if it is enable or disabled
   if ($lockdown) {
     $lockdown_status = "Enabled"  
   }   
   else {
     $lockdown_status = "Disabled"
   }


 ####################################################################
 ## Output to file 
 $filename = $esxi+"_MHC_Results.txt"
 $Output_MHC = ".\Results\MHC_Info\$filename"

 if(!(Test-Path $Output_MHC)) {
     New-Item -ItemType "file" -Path $Output_MHC
  }

 else {
     Clear-Content $Output_MHC
  }
  Add-Content -Path $Output_MHC -Value "## Number of failed attemps before account is locked : $acclock" 
  Add-Content -Path $Output_MHC -Value "## Seconds of lock duration after lock account : $timelock" 
  Add-Content -Path $Output_MHC -Value "## Password Complexity  : $complexity"  
  Add-Content -Path $Output_MHC -Value "## Log rotation is set to rotate $rotate files with $size size each. Total size $total" 
  Add-Content -Path $Output_MHC -Value "## Remote host sents logs to : $syslogglobal_set"  
  Add-Content -Path $Output_MHC -Value "## On esxi $esxi the following ports are open:"
  Add-Content -Path $Output_MHC -Value  (Get-VMHostFirewallException -VMHost $esxi -Enabled:$True | select Name,Enabled| out-string)
  Add-Content -Path $Output_MHC -Value "## SNMP Service is in status : $smtp_st"
  Add-Content -Path $Output_MHC -Value "## Policy is set to : $smtp_po "
  Add-Content -Path $Output_MHC -Value "## Remote Tech Support Service is in status : $tsm_st"
  Add-Content -Path $Output_MHC -Value "## Policy is set to : $tsm_po "
  Add-Content -Path $Output_MHC -Value "## Local Tech Support Service is in status : $tsmssh_st"
  Add-Content -Path $Output_MHC -Value "## Policy is set to : $tsmssh_po "
  Add-Content -Path $Output_MHC -Value "## NTP Serive is in status : $ntp_st"
  Add-Content -Path $Output_MHC -Value "## Policy is set to : $ntp_po "
  Add-Content -Path $Output_MHC -Value "## DCUI Serive is in status : $dcui_st"
  Add-Content -Path $Output_MHC -Value "## Policy is set to : $dcui_po "
  Add-Content -Path $Output_MHC -Value "## DCUI Timeout Value : $dcui_timer"
  Add-Content -Path $Output_MHC -Value "## Lockdown mode is : $lockdown_status"
  Add-Content -Path $Output_MHC -Value " "

 ####################################################################  
 ####################################################################


 # vSwitch and port security policy.

 foreach ($esxi in Get-VMHost){
     foreach($vSwitch in $esxi | Get-VirtualSwitch -Standard){
       Add-Content -Path $Output_MHC -Value $vSwitch.Name
       Add-Content -Path $Output_MHC -Value "`t Promiscuous mode enabled: $($vSwitch.ExtensionData.Spec.Policy.Security.AllowPromiscuous)"
       Add-Content -Path $Output_MHC -Value "`t Forged transmits enabled: $($vSwitch.ExtensionData.Spec.Policy.Security.ForgedTransmits)"
       Add-Content -Path $Output_MHC -Value "`t MAC Changes enabled.....: $($vSwitch.ExtensionData.Spec.Policy.Security.MacChanges)"
       Add-Content -Path $Output_MHC -Value ""

        foreach($portgroup in ($esxi.ExtensionData.Config.Network.Portgroup | where {$_.Vswitch -eq $vSwitch.Key})){
          Add-Content -Path $Output_MHC $portgroup.Spec.Name       
          Add-Content -Path $Output_MHC -Value "`t Promiscuous mode enabled: $(
            If ($portgroup.Spec.Policy.Security.AllowPromiscuous -eq $null) {Write-Output $vSwitch.ExtensionData.Spec.Policy.Security.AllowPromiscuous } Else {Write-Output $portgroup.Spec.Policy.Security.AllowPromiscuous })"
          Add-Content -Path $Output_MHC -Value "`t Forged transmits enabled: $(
            If ($portgroup.Spec.Policy.Security.ForgedTransmits -eq $null) {Write-Output $vSwitch.ExtensionData.Spec.Policy.Security.ForgedTransmits } Else {Write-Output $portgroup.Spec.Policy.Security.ForgedTransmits })"
         Add-Content -Path $Output_MHC -Value "`t MAC Changes enabled.....: $( 
            If ($portgroup.Spec.Policy.Security.MacChanges -eq $null) {Write-Output $vSwitch.ExtensionData.Spec.Policy.Security.MacChanges } Else {Write-Output $portgroup.Spec.Policy.Security.MacChanges })"
        Add-Content -Path $Output_MHC -Value ""
        }
         
    }
    foreach($vSwitch in $esxi | Get-VirtualSwitch -Distributed){
       Add-Content -Path $Output_MHC -Value $vSwitch.Name
       Add-Content -Path $Output_MHC -Value "Promiscuous mode enabled: $($vSwitch.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.AllowPromiscuous.Value)"
       Add-Content -Path $Output_MHC -Value "Forged transmits enabled: $($vSwitch.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.ForgedTransmits.Value)"
       Add-Content -Path $Output_MHC -Value "MAC Changes enabled.....: $($vSwitch.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.MacChanges.Value)"
       Add-Content -Path $Output_MHC -Value ""

        foreach($portgroup in (Get-VirtualPortGroup -Distributed -VirtualSwitch $vSwitch)){
           Add-Content -Path $Output_MHC -Value "`n`t`t"$portgroup.Name
           Add-Content -Path $Output_MHC -Value "`t Promiscuous mode enabled: $($portgroup.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.AllowPromiscuous.Value)"
           Add-Content -Path $Output_MHC -Value "`t Forged transmits enabled: $($portgroup.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.ForgedTransmits.Value)"
           Add-Content -Path $Output_MHC -Value "`t MAC Changes enabled.....: $($portgroup.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.MacChanges.Value)"
           Add-Content -Path $Output_MHC -Value ""  
        }
    }
} 

 ####################################################################  
 ####################################################################
 # Fill in excel file



$ExcelWorkBook = $Excel.Workbooks.Open($excel_file_path)
$Excelfirsttab = $Excel.WorkSheets.item("Doc Control")
$ExcelWorkSheet = $Excel.WorkSheets.item("Tech Spec")
$Excelfirsttab.activate()

## Write esxi info
 $Excelfirsttab.Cells.Item(36,3) = "MHC on esxi "+ $esxi+" Performed on "+$date
 $ExcelWorkSheet.activate()

 ## Check Password length
 $ExcelWorksheet.Cells.Item(3,9).Interior.ColorIndex = 6
 $ExcelWorkSheet.Cells.Item(3,9) = "Warning - Check Password complexity"
 

 ## Check password complexity
if ($Init_PSW_CMPL -ne $complexity) {
 $ExcelWorksheet.Cells.Item(4,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(4,9) = "KO - The current value is "+$complexity
 }
 else {
   $ExcelWorkSheet.Cells.Item(4,9) = "OK - The current value is "+$complexity
 }


 ## Check Seconds of lock duration after lock account 
if ($Init_Lock_Duration -ne $timelock) {
 $ExcelWorksheet.Cells.Item(5,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(5,9) = "KO - The current value is "+$timelock
 }
 else {
   $ExcelWorkSheet.Cells.Item(5,9) = "OK - The current value is "+$timelock
 }


 ## Check failed attemps before account is locked
if ($Init_PSW_Length -ne $acclock) {
 $ExcelWorksheet.Cells.Item(6,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(6,9) = "KO - The current value is "+$acclock
 }else {
   $ExcelWorkSheet.Cells.Item(6,9) = "OK - The current value is "+$acclock
 }
 

## Check Log files
  $ExcelWorkSheet.Cells.Item(7,9) = "OK - It is set by default"
  $ExcelWorkSheet.Cells.Item(8,9) = "OK - It is set by default"
  $ExcelWorkSheet.Cells.Item(9,9) = "OK - It is set by default"
  $ExcelWorkSheet.Cells.Item(10,9) = "OK - It is set by default"
  

   ## Check failed attemps before account is locked
if ($syslogglobal -eq "") {
 $ExcelWorksheet.Cells.Item(11,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(11,9) = "K.O. - Remote logging is not set"
 }
 else {
    $ExcelWorkSheet.Cells.Item(11,9) = "OK - It is set to  "+$syslogglobal_set
 }
 

 ## Check Vlan and vm sharing
  $ExcelWorkSheet.Cells.Item(13,9) = "OK - It is not shared by default"
  $ExcelWorkSheet.Cells.Item(14,9) = "OK - It is not shared by default"
  $ExcelWorkSheet.Cells.Item(18,9) = "OK - It is not shared by default"
  $ExcelWorkSheet.Cells.Item(19,9) = "OK - It is not shared by default"
  $ExcelWorkSheet.Cells.Item(20,9) = "OK - It is not shared by default"
  $ExcelWorkSheet.Cells.Item(21,9) = "OK - It is not shared by default"
  
  
  #  Check Open Ports
 $portshc = (Get-VMHostFirewallException -VMHost $esxi -Enabled:$True | select Name | out-string)
 $ExcelWorksheet.Cells.Item(22,9).Interior.ColorIndex = 6
 $ExcelWorkSheet.Cells.Item(22,9) = "Warning - Following ports are open `n"+$portshc

 #Check smtp 
 if ($Init_Stopped -ne $smtp_st) {
 $ExcelWorksheet.Cells.Item(23,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(23,9) = "KO - SMTP is "+$smtp_st
 $ExcelWorkSheet.Cells.Item(24,9) = "KO - Community name public exist"
 $ExcelWorkSheet.Cells.Item(25,9) = "KO - Community name private exist"
 }else {
 $ExcelWorkSheet.Cells.Item(23,9) = "OK - SMTP is "+$smtp_st
 $ExcelWorkSheet.Cells.Item(24,9) = "OK - Community name public does not exist"
 $ExcelWorkSheet.Cells.Item(25,9) = "OK - Community name private does not exist"
 }
 
 
## Check SSH
if ($Init_Stopped -ne $tsm_st) {
 $ExcelWorksheet.Cells.Item(26,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(26,9) = "K.O. - Service is"+$tsm_st+" and policy is set to "+$tsm_po
 }
 else {
    $ExcelWorkSheet.Cells.Item(26,9) = "OK - Service is "+$tsm_st+" and policy is set to "+$tsm_po
 }
 

## Check Default local console
if ($Init_Stopped -ne $tsmssh_st) {
 $ExcelWorksheet.Cells.Item(27,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(27,9) = "K.O. - Service is"+$tsmssh_st
 }
 else {
    $ExcelWorkSheet.Cells.Item(27,9) = "OK - Service is "+$tsmssh_st
 }
 

 ## Check NTP
if ($Init_Running -ne $ntp_st -and $Init_policy -ne $ntp_po) {
 $ExcelWorksheet.Cells.Item(28,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(28,9) = "K.O. - Service is"+$ntp_st+" and policy is set to "+$ntp_po
 }
 else {
    $ExcelWorkSheet.Cells.Item(28,9) = "OK - Service is "+$ntp_st+" and policy is set to "+$ntp_po
 }
 

  ## Check DCUI
if ($Init_Running -ne $dcui_st -and $Init_policy -ne $dcui_po) {
 $ExcelWorksheet.Cells.Item(29,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(29,9) = "K.O. - Service is"+$dcui_st+" and policy is set to "+$dcui_po
 }
 else {
    $ExcelWorkSheet.Cells.Item(29,9) = "OK - Service is "+$dcui_st+" and policy is set to "+$dcui_po
 }
 

   ## Check DCUI lockdown
if ($Init_timmer -ne $dcui_timer) {
 $ExcelWorksheet.Cells.Item(30,9).Interior.ColorIndex = 3
 $ExcelWorkSheet.Cells.Item(30,9) = "K.O. - It is set to "+$dcui_timer
 }
 else {
    $ExcelWorkSheet.Cells.Item(30,9) = "OK - It is set to "+$dcui_timer
 }
 
 ## Check Vlan and vm sharing
  $ExcelWorkSheet.Cells.Item(32,9) = "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(33,9) = "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(34,9) = "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(35,9) =  "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(36,9) =  "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(37,9) =  "OK - It is set as required by default"
  $ExcelWorkSheet.Cells.Item(38,9) =  "OK - It is set as required by default"

  ##### Save file
   
  $xlspath = 'c:\Scripts\MHC\Results\MHC_Excel\'+$esxi+'_MHC.xls'
  $ExcelWorkBook.SaveAs($xlspath,[Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
  $Excel.quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    }
  }

  #####################   
  ## Zip results
   $Date = (Get-Date -f "ddMMyyyy")
  #   $source = ".\Results"
   #  $destination = ".\"+$Date+".zip",
   #[io.compression.zipfile]::CreateFromDirectory($Source, $destination)

 $compress = @{
Path=".\Results"
CompressionLevel = "Fastest"
DestinationPath = $destination+"_"+$Date+".zip"
}
$sourcefile = $destination+"_"+$Date+".zip"
$Archive = ".\Archive"
Compress-Archive @compress
Move-Item -Path $sourcefile -Destination $Archive
Remove-Item -path $destination -recurse
