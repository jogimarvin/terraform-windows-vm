# Variables
$domain = "rocku.com"
$domainUser = "rocku\azureadmin"
$domainPassword = ConvertTo-SecureString "SuperSecretPass123!" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($domainUser, $domainPassword)

# Join the computer to the domain
Add-Computer -DomainName $domain -Credential $credential -Restart -Force
