$emailSubject = "PCI Data Loss Prevention Test Email"
$fakeCreditCard = "MasterCard 5555555555554444 "

function Exit-ThisScript()
{
    Write-Output "Disconnecting and exiting..."
    Get-PSSession | Remove-PSSession
    Disconnect-AzureAD
    Exit
}

$azureConnection = Connect-AzureAD
Connect-IPPSSession -UserPrincipalName $azureConnection.Account

$numberOfCreditCards = (Get-DlpComplianceRule | Where-Object {$_.IsValid -and $_.AccessScope -eq 'InOrganization' -and $_.Mode -eq 'Enforce' -and !$_.Disabled -and $_.BlockAccess -and $_.ContentContainsSensitiveInformation.name -eq 'Credit Card Number'}).ContentContainsSensitiveInformation.mincount | Select-Object -First 1

if (-not ($numberOfCreditCards -ge 1))
{
    Write-Warning "Unable to find a credit card rule to test."
    Exit-ThisScript
}

$sb = [System.Text.StringBuilder]::new()
foreach( $i in 1..$numberOfCreditCards)
{
    [void]$sb.AppendLine( $fakeCreditCard )
}
$emailBody = $sb.ToString()

$emailAccountId = $azureConnection.Account.Id

Start-Process "mailto:${emailAccountId}?Subject=$emailSubject&Body=$emailBody"
Exit-ThisScript