<#
.SYNOPSIS
  This script will upgrade the existing SCOM agent to the latest version (While keeping it's existing settings and also add a new SCOM instance). It will also do the following:
  - Add the system to Azure Patching System
  - Add the system to Defender ATP
 
.INPUTS
  This runbook has no inputs.

.OUTPUTS
  None

.NOTES
  Version:        1.0
  Author:         Robert Curtiss
  Creation Date:  2021-3-24
  Purpose/Change: Initial runbook development
  
 #> 


#Upgrade existing agent 
Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList '/i', '"\\servername\newmomagent\MOMAgent.msi"','AcceptEndUserLicenseAgreement=1', '/qn' -Wait

#Variables
$ManagementServer = "servername.yourcompany.com"
$MGMTGroupName = "MANAGEMENT GROUP NAME"
$Port = "5723"

#Add New Managment Group 
$Agent = New-Object -ComObject AgentConfigManager.MgmtSvcCfg
$Agent.AddManagementGroup("$MGMTGroupName", "$ManagementServer", "$Port")

#Add OMS Workspace ID (used for Windows updates) 
$workspaceId = "INSERT WORKSPACE ID HERE"
$workspaceKey = "INSERT WORKSPACE KEY HERE"
$mma = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'
$mma.AddCloudWorkspace($workspaceId, $workspaceKey)
$mma.ReloadConfiguration()

#Add OMS Workspace ID (used for Defender ATP) 
$workspaceId = "INSERT WORKSPACE ID HERE"
$workspaceKey = "INSERT WORKSPACE KEY HERE"
$mma = New-Object -ComObject 'AgentConfigManager.MgmtSvcCfg'
$mma.AddCloudWorkspace($workspaceId, $workspaceKey)
$mma.ReloadConfiguration()