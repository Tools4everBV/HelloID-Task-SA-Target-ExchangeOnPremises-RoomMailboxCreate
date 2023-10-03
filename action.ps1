# HelloID-Task-SA-Target-ExchangeOnPremises-RoomMailboxCreate
#############################################################
# Form mapping
$formObject = @{
    Name               = $form.RoomName
    DisplayName        = $form.RoomName
    PrimarySmtpAddress = $form.PrimarySmtpAddress
    OrganizationalUnit = $form.OrganizationalUnit
    ResourceCapacity   = $form.ResourceCapacity
    Password           = (ConvertTo-SecureString -AsPlainText $form.password -Force)
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnPremises action: [RoomMailboxCreate] for: [$($formObject.DisplayName)]"
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'New-Mailbox'
    $IsConnected = $true

    $mailbox = New-Mailbox @formObject -Room -ErrorAction Stop

    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $mailbox.ExchangeGuid
        TargetDisplayName = $formObject.DisplayName
        Message           = "ExchangeOnPremises action: [RoomMailboxCreate] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnPremises action: [RoomMailboxCreate] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.DisplayName
        TargetDisplayName = $formObject.DisplayName
        Message           = "Could not execute ExchangeOnPremises action: [RoomMailboxCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [RoomMailboxCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
#############################################################
