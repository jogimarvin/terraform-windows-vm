# Variables
$domain = "rocku.com"
$domainUser = "rocku\azureadmin"
$domainPassword = ConvertTo-SecureString "SuperSecretPass123!" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($domainUser, $domainPassword)

# Join the computer to the domain
try {
    Add-Computer -DomainName $domain -Credential $credential -Restart -Force
    Write-Output "Domain join successful."
} catch {
    Write-Error "Domain join failed: $_"
}
