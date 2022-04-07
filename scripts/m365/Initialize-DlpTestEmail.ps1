# Script for generating test emails for triggering data loss prevention (DLP) rules

# A hash table of all rules this script is capable of testing,
# each line is in the following format:
#   'Rulle to Check' = 'Example string to trigger the rule'
$rulesToCheck = @{
    'Credit Card Number' = 'MasterCard 5555555555554444'
    'U.S. Social Security Number (SSN)' = 'SSN 694-09-5553'
    'U.S. Individual Taxpayer Identification Number (ITIN)' = 'itin 957-82-4338'
    'U.S. Bank Account Number' = 'Debit Account 021000021'
    'ABA Routing Number' = 'aba 121042882'
    'U.S. / U.K. Passport Number' = 'passport no P31195855 17 Sep 2031'
    'Drug Enforcement Agency (DEA) number' = 'dea KV2993548'
    'International Classification of Diseases (ICD-9-CM)' = 'ICD-9-CM Diagnosis Code 425.11 Hypertrophic obstructive cardiomyopathy'
    'International Classification of Diseases (ICD-10-CM)' = 'ICD-10-CM Diagnosis Code A15.7 Primary respiratory tuberculosis'
}

$global:externalEmailAddress = $null
$emailSubject = "Data Loss Prevention (DLP) Test Email"

# Install required modules if missing
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

# Return the number of credit cards required to trigger the rule
function Get-MinCount {
    param(
        [Parameter(Mandatory = $true)] [string]$chosenAccessScope,
        [Parameter(Mandatory = $true)] [string]$selectedRule)
    ((Get-DlpComplianceRule | Where-Object { $_.IsValid -and $_.AccessScope -eq $chosenAccessScope -and $_.Mode -eq 'Enforce' -and !$_.Disabled  }).ContentContainsSensitiveInformation | Where-Object {$_.name -eq $selectedRule}).mincount | Sort-Object -Descending { [int]$_ } | Select-Object -First 1
}

# Pick between 'InOrganization' and 'NotInOrganization' test
function Get-WhichModeToTest {
    $internal = New-Object System.Management.Automation.Host.ChoiceDescription '&Internal', 'Generate internal test email'
    $external = New-Object System.Management.Automation.Host.ChoiceDescription '&External', 'Generate external test email'
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($internal, $external
    )

    $title = "We detected both internal and external $selectedRule rules."
    $message = 'For which rule do you want to generate a test email?'
    $result = $host.ui.PromptForChoice($title, $message, $options, 0)

    return $result
}

# Ask to enter an email address for 'NotInOrganization' test
function Get-ExternalEmail {
    if ($null -eq $global:externalEmailAddress){
        $global:externalEmailAddress = Read-Host -Prompt 'Please enter an email address outside of your organization to send the test email to'
    }
    return $global:externalEmailAddress
}

# Generate test email
function Send-Email {
    param(
        [Parameter(Mandatory = $true)] [string]$emailAddress,
        [Parameter(Mandatory = $true)] [int]$numberOfCRepetitions,
        [Parameter(Mandatory = $true)] [string]$fakeSensetiveString
        )

    $sb = [System.Text.StringBuilder]::new()
    foreach ( $i in 1..$numberOfCRepetitions) {
        [void]$sb.AppendLine( $fakeSensetiveString )
        [void]$sb.AppendLine( ' ' )
    }
    $emailBody = $sb.ToString()

    Start-Process "mailto:${emailAddress}?Subject=$emailSubject&Body=$emailBody"
    Write-Output "Please review and finish sending the test email in your default email app."
}

function Start-Test {
    $selectedRule = $rulesToCheck.Keys | Out-Gridview -Title "Select your choice" -OutputMode Single
    if (-not $selectedRule) { 
        Write-Host "No rule selected."
        Exit-ThisScript
    }

    Write-Host "You chose to test the '$selectedRule' rule. Retrieving the rule configurations..." 

    $numberOfInternal = Get-MinCount 'InOrganization' $selectedRule
    $numberOfExternal = Get-MinCount 'NotInOrganization' $selectedRule
    $fakeSensetiveString =  $rulesToCheck[$selectedRule]

    if (-not ($numberOfInternal -ge 1 -or $numberOfExternal -ge 1)) {
        Write-Warning "Unable to find $selectedRule rule to test. Please create or enable a DLP policy with $selectedRule rule enforcement."
        Exit-ThisScript
    }

    if ($numberOfInternal -ge 1 -and $numberOfExternal  -ge 1) {
        switch (Get-WhichModeToTest) {
            0 { Send-Email $azureConnection.Account.Id $numberOfInternal $fakeSensetiveString }   # internal test
            1 { Send-Email (Get-ExternalEmail) $numberOfExternal $fakeSensetiveString }           # external test
        }
        Exit-ThisScript
    }

    if ($numberOfInternal -ge 1) {
        Send-Email $azureConnection.Account.Id $numberOfInternal $fakeSensetiveString             # internal test
        Exit-ThisScript
    }

    Send-Email (Get-ExternalEmail) $numberOfExternal $fakeSensetiveString                         # external test
    Exit-ThisScript
}

# Close open connections and exit
function Exit-ThisScript() {
    $finished = Read-Host "Exit? [y/n]"
    if ($finished -eq "n"){
        Start-Test                                  # repeat until the user chooses to exit
        return
    } 
    Write-Output "Disconnecting and exiting..."
    Get-PSSession | Remove-PSSession
    Disconnect-AzureAD
    Exit
}

Install-Prerequisites
Write-Output "Logging in to Azure AD. The login window may be behind other windows..."
$azureConnection = Connect-AzureAD
Connect-IPPSSession -UserPrincipalName $azureConnection.Account

Start-Test