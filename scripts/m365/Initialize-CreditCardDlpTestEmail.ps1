$emailSubject = "PCI Data Loss Prevention Test Email"
$fakeCreditCard = "MasterCard 5555555555554444 "

# install required modules if missing
function Install-Prerequisites {
    function Install-PrerequisiteModule {
        param([Parameter(Mandatory = $true)] $moduleName)    
        if (Get-Module -ListAvailable -Name $moduleName) {
            Write-Host "$moduleName already installed"
        } 
        else {
            try {
                Write-Host "Installing missing module: $moduleName..."
                Install-Module -Name $moduleName
            }
            catch [Exception] {
                $_.message 
                Exit
            }
        }
    }

    Install-PrerequisiteModule AzureAD
    Install-PrerequisiteModule ExchangeOnlineManagement
}

# return the number of credit cards required to trigger the rule
function Get-NumberOfCreditCards {
    param([Parameter(Mandatory = $true)] [string]$chosenAccessScope)
    (Get-DlpComplianceRule | Where-Object { $_.IsValid -and $_.AccessScope -eq $chosenAccessScope -and $_.Mode -eq 'Enforce' -and !$_.Disabled -and $_.ContentContainsSensitiveInformation.name -eq 'Credit Card Number' -and $_.ReportSeverityLevel -eq 'High' }).ContentContainsSensitiveInformation.mincount | Select-Object -First 1
}

# pick between 'InOrganization' and 'NotInOrganization' test
function Get-WhichModeToTest {
    $internal = New-Object System.Management.Automation.Host.ChoiceDescription '&Internal', 'Generate internal test email'
    $external = New-Object System.Management.Automation.Host.ChoiceDescription '&External', 'Generate external test email'
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($internal, $external)

    $title = 'We detected both internal and external Credit Card Number rules.'
    $message = 'For which rule do you want to generate a test email?'
    $result = $host.ui.PromptForChoice($title, $message, $options, 0)

    return $result
}

# ask to enter an email address for 'NotInOrganization' test
function Get-ExternalEmail {
    Read-Host -Prompt 'Please enter an email address outside of your organization to send the test email to'
}

# generate test email
function Send-Email {
    param(
        [Parameter(Mandatory = $true)] [string]$emailAddress,
        [Parameter(Mandatory = $true)] [int]$numberOfCreditCards)

    $sb = [System.Text.StringBuilder]::new()
    foreach ( $i in 1..$numberOfCreditCards) {
        [void]$sb.AppendLine( $fakeCreditCard )
    }
    $emailBody = $sb.ToString()

    Start-Process "mailto:${emailAddress}?Subject=$emailSubject&Body=$emailBody"
    Write-Output "Please review and finish sending the test email in your default email app."
}

# close open connections and exit
function Exit-ThisScript() {
    Write-Output "Disconnecting and exiting..."
    Get-PSSession | Remove-PSSession
    Disconnect-AzureAD
    Exit
}

Install-Prerequisites
Write-Output "Logging in to Azure AD. The login window may be behind other windows..."
$azureConnection = Connect-AzureAD
Connect-IPPSSession -UserPrincipalName $azureConnection.Account

$numberOfInternalCards = Get-NumberOfCreditCards('InOrganization')
$numberOfExternalCards = Get-NumberOfCreditCards('NotInOrganization')

if (-not ($numberOfInternalCards -ge 1 -or $numberOfExternalCards -ge 1)) {
    Write-Warning "Unable to find a credit card rule to test. Please create or enable PCI DLP Policy."
    Exit-ThisScript
}

if ($numberOfInternalCards -ge 1 -and $numberOfExternalCards -ge 1) {
    switch (Get-WhichModeToTest) {
        0 { Send-Email $azureConnection.Account.Id $numberOfInternalCards }   # internal test
        1 { Send-Email (Get-ExternalEmail) $numberOfExternalCards }             # external test
    }
    Exit-ThisScript
}

if ($numberOfInternalCards -ge 1) {
    Send-Email $azureConnection.Account.Id $numberOfInternalCards              # internal test
    Exit-ThisScript
}

Send-Email (Get-ExternalEmail) $numberOfExternalCards                            # external test
Exit-ThisScript