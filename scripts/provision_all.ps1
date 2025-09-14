# === Enable WinRM (Only if not already running) ===
if ((Get-Service winrm).Status -ne 'Running') {
    Write-Output "Enabling WinRM..."
    winrm quickconfig -quiet
    winrm set winrm/config/service/auth '@{Basic="true"}'
    winrm set winrm/config/service '@{AllowUnencrypted="true"}'
    Enable-PSRemoting -Force
    Set-NetFirewallRule -Name "WINRM-HTTP-In-TCP-PUBLIC" -RemoteAddress Any
}

# === Install AD DS (Only if not installed) ===
$adInstalled = Get-WindowsFeature AD-Domain-Services
if (-not $adInstalled.Installed) {
    Write-Output "Installing AD DS..."
    Install-WindowsFeature AD-Domain-Services -IncludeManagementTools
}

# === Promote to Domain Controller (Only if not already promoted) ===
try {
    $domain = Get-ADDomain -ErrorAction Stop
    Write-Output "Domain already exists: $($domain.Name)"
} catch {
    Write-Output "Promoting server to Domain Controller..."
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

    # Import AD module to create OUs
    Import-Module ActiveDirectory

    # List of OUs to create
    $ousToCreate = @("Sales", "HR", "IT", "Finance")

    foreach ($ouName in $ousToCreate) {
        $ouDn = "OU=$ouName,DC=rocku,DC=com"
        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouDn'" -ErrorAction SilentlyContinue)) {
            New-ADOrganizationalUnit -Name $ouName -Path "DC=rocku,DC=com"
            Write-Output "Created OU: $ouName"
        } else {
            Write-Output "OU $ouName already exists, skipping creation."
        }
    }

    exit  # Server will reboot after domain promotion
}

# === Set up logging ===
$logPath = "C:\Logs\user_provisioning.log"
if (-not (Test-Path "C:\Logs")) { New-Item -Path "C:\Logs" -ItemType Directory -Force }

function Log($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $msg" | Out-File -FilePath $logPath -Append
}

# === User Lifecycle Provisioning ===
$excelPath = ".\user.xlsx"
if (Test-Path $excelPath) {
    try {
        Write-Output "Starting user provisioning..."
        Log "Starting user provisioning from $excelPath"

        Import-Module ActiveDirectory
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($excelPath)
        $sheet = $workbook.Sheets.Item(1)

        $row = 2
# ...[previous code above remains unchanged]...
        
        while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
            $action       = $sheet.Cells.Item($row, 1).Value()
            $username     = $sheet.Cells.Item($row, 2).Value()
            $firstName    = $sheet.Cells.Item($row, 3).Value()
            $lastName     = $sheet.Cells.Item($row, 4).Value()
            $password     = $sheet.Cells.Item($row, 5).Value()
            $department   = $sheet.Cells.Item($row, 6).Value()
            $newDepartment= $sheet.Cells.Item($row, 7).Value()  # New column for mover
        
            $fullName = "$firstName $lastName"
            $userPrincipalName = "$username@rocku.com"
        
            switch ($action) {
                "Joiner" {
                    Log "🧪 Processing Joiner for user $username, department: $department"
                
                    if (-not $department) {
                        Log "⚠️ No department for $username. Skipping."
                        $row++
                        continue
                    }
                
                    $ou = "OU=$department,DC=rocku,DC=com"
                
                    $existingUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
                    if ($existingUser) {
                        Write-Warning "User $username already exists. Skipping Joiner action."
                        Log "Skipped Joiner: User $username already exists"
                        $row++
                        continue
                    }
                
                    $ouExists = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ou'" -ErrorAction SilentlyContinue
                    if (-not $ouExists) {
                        Write-Warning "⛔ OU '$ou' does not exist. Skipping user $username."
                        Log "⛔ OU does not exist for user $username. OU string: $ou"
                        $row++
                        continue
                    }
                
                    # Proceed to create user
                    try {
                        New-ADUser `
                            -Name "$firstName $lastName" `
                            -SamAccountName $username `
                            -UserPrincipalName "$username@rocku.com" `
                            -GivenName $firstName `
                            -Surname $lastName `
                            -Path $ou `
                            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                            -Enabled $true
                
                        Write-Output "✅ Created user: $username"
                        Log "✅ Created user: $username in OU: $ou"
                    } catch {
                        Write-Warning "❌ Failed to create user $username: $_"
                        Log "❌ Failed to create user $username: $_"
                    }
                }

                "Mover" {
                    $ou = "OU=$newDepartment,DC=rocku,DC=com"
        
                    try {
                        $userObj = Get-ADUser -Filter "SamAccountName -eq '$username'" -Properties DistinguishedName
        
                        # Validate target OU exists
                        if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ou'" -ErrorAction SilentlyContinue)) {
                            Write-Warning "Target OU '$ou' does not exist. Skipping mover for user $username."
                            Log "Mover skipped. OU does not exist for $username: $ou"
                            $row++
                            continue
                        }
        
                        Move-ADObject -Identity $userObj.DistinguishedName -TargetPath $ou
                        Write-Output "Moved user: $username to $ou"
                        Log "Moved user: $username to $ou"
                    } catch {
                        Write-Warning "Failed to move user $username. $_"
                        Log "Failed to move user $username: $_"
                    }
                }
        
                "Leaver" {
                    try {
                        Disable-ADAccount -Identity $username
                        Write-Output " Disabled user: $username"
                        Log "Disabled user: $username"
                    } catch {
                        Write-Warning "Failed to disable user $username. $_"
                        Log "Failed to disable user $username: $_"
                    }
                }
        
                default {
                    Write-Warning "Unknown action '$action' for user $username. Skipping."
                    Log "Skipped unknown action '$action' for $username"
                }
            }
        
            $row++
        }

    } catch {
        Write-Error "Error processing Excel file: $_"
        Log "Error processing Excel file: $_"
    } finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
} else {
    Write-Warning "Excel file not found at: $excelPath"
    Log "Excel file not found: $excelPath"
}
