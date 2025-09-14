Import-Module ActiveDirectory

$users = Import-Csv "user_lifecycle.csv"

foreach ($user in $users) {
    switch ($user.Action) {
        "Joiner" {
            New-ADUser `
                -Name "$($user.FirstName) $($user.LastName)" `
                -SamAccountName $user.Username `
                -UserPrincipalName "$($user.Username)@rocku.com" `
                -GivenName $user.FirstName `
                -Surname $user.LastName `
                -Path $user.OU `
                -AccountPassword (ConvertTo-SecureString $user.TempPassword -AsPlainText -Force) `
                -Enabled $true
            Write-Output "Created user: $($user.Username)"
        }

        "Mover" {
            Move-ADObject `
                -Identity "CN=$($user.FirstName) $($user.LastName),$user.OU" `
                -TargetPath $user.NewOU
            Write-Output "Moved user: $($user.Username)"
        }

        "Leaver" {
            Disable-ADAccount -Identity $user.Username
            Write-Output "Disabled user: $($user.Username)"
        }
    }
}
