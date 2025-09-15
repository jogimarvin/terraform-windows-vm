# ===============================
# Stage 1: Install AD DS and Promote Server
# ===============================

# Install AD DS Role
Install-WindowsFeature AD-Domain-Services -IncludeManagementTools

# Import ADDSDeployment module
Import-Module ADDSDeployment

# Promote to domain controller
Install-ADDSForest `
  -DomainName "rocku.com" `
  -CreateDnsDelegation:$false `
  -DatabasePath "C:\Windows\NTDS" `
  -DomainMode "7" `
  -ForestMode "7" `
  -InstallDns:$true `
  -LogPath "C:\Windows\NTDS" `
  -SysvolPath "C:\Windows\SYSVOL" `
  -Force `
  -SafeModeAdministratorPassword (ConvertTo-SecureString "AnotherSecretPass123!" -AsPlainText -Force)

# Note: Install-ADDSForest will automatically reboot the server after promotion

# ===============================
# Download Stage 2 script for manual execution after reboot
# ===============================
$stage2Url = "https://rockuadscripts.blob.core.windows.net/scripts/stage2-setup.ps1"
$stage2Path = "C:\Scripts\stage2-setup.ps1"

# Ensure directory exists
New-Item -Path "C:\Scripts" -ItemType Directory -Force

# Download the Stage 2 script
Invoke-WebRequest -Uri $stage2Url -OutFile $stage2Path

Write-Host "Stage 2 script downloaded to $stage2Path. You can run it manually after login."
