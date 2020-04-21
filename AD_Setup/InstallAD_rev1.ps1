Param(
[string]$DomainName='Ronsulting.net',
[string]$DomainNETBIOSName='Ronsulting',
[string]$Password='Secret321!'
)

#Import-module ADDSDeployment

 

# Add Domain Services to Server

INSTALL-WindowsFeature AD-Domain-Services -IncludeManagementTools -IncludeAllSubFeature

 

# Reset local Administrator password to

# One defined in Script

NET USER Administrator $Password

 

# Install our Shiny new Forest. It's a good place for SQuirreLs to see *'s

#
Import-Module ADDSDeployment
Install-ADDSForest `
-CreateDnsDelegation:$false `
-DatabasePath "C:\Windows\NTDS" `
-DomainMode "Win2012R2" `
-DomainName $DomainName `
-DomainNetbiosName $DomainNETBIOSName `
-ForestMode "Win2012R2" `
-InstallDns:$true `
-LogPath "C:\Windows\NTDS" `
-NoRebootOnCompletion:$false `
-SysvolPath "C:\Windows\SYSVOL" `
-SafeModeAdministratorPassword (CONVERTTO-SecureString $Password -asplaintext -force) `
-Force:$true
 

# Restart and Get 'er done

RESTART-COMPUTER -force
