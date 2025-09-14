# === Enable WinRM (Only if not already running) ===
if ((Get-Service winrm).Status -ne 'Running') {
    winrm quickconfig -quiet
    winrm set winrm/config/service/auth '@{Basic="true"}'
    winrm set winrm/config/service '@{AllowUnencrypted="true"}'
    Enable-PSRemoting -Force
    Set-NetFirewallRule -Name "WINRM-HTTP-In-TCP-PUBLIC" -RemoteAddress Any
}

# === Install AD DS (Only if not installed) ===
$adInstalled = Get-WindowsFeature AD-Domain-Services
if ($adInstalled.Installed -eq $false) {
    Install-WindowsFeature AD-Domain-Services -IncludeManagementTools
}

# === Promote to Domain Controller (Only if not already promoted) ===
try {
    $domain = Get-ADDomain -ErrorAction Stop
} catch {
    Import-Module ADDSDeployment
    Install-ADDSForest `
        -DomainName "rocku.com" `
        -CreateDnsDelegation:$false `
        -DatabasePath "C:\Windows\NTDS" `
        -DomainMode "WinThreshold" `
        -DomainNetbiosName "ROCKU" `
        -ForestMode "WinThreshold" `
        -InstallDns:$true `
        -LogPath "C:\Windows\NTDS" `
        -SysvolPath "C:\Windows\SYSVOL" `
        -Force:$true `
        -SafeModeAdministratorPassword (ConvertTo-SecureString "P@ssword1234!" -AsPlainText -Force)

    # This triggers a reboot — so exit the script after promotion
    exit
}

# === User Lifecycle Provisioning ===
$excelPath = ".\users.xlsx"
if (Test-Path $excelPath) {
    try {
        Import-Module ActiveDirectory
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($excelPath)
        $sheet = $workbook.Sheets.Item(1)

        $row = 2
        while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
            $action = $sheet.Cells.Item($row, 1).Value()
            $username = $sheet.Cells.Item($row, 2).Value()
            $fullname = $sheet.Cells.Item($row, 3).Value()
            $ou = "OU=Users,DC=rocku,DC=com"

            if ($action -eq "Joiner" -and -not (Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue)) {
                Write-Output "Creating user: $username"
                New-ADUser -Name $fullname -SamAccountName $username -AccountPassword (ConvertTo-SecureString "P@ssword123" -AsPlainText -Force) -Enabled $true -Path $ou
            } elseif ($action -eq "Mover") {
                Write-Output "Updating user: $username (Mover)"
                Set-ADUser -Identity $username -Department "NewDept"
            } elseif ($action -eq "Leaver") {
                Write-Output "Disabling user: $username"
                Disable-ADAccount -Identity $username
            }

            $row++
        }

        $workbook.Close($false)
        $excel.Quit()
    } catch {
        Write-Error "Error processing users.xlsx: $_"
    }
}
