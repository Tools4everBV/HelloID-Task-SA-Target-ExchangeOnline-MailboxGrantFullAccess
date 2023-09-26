# HelloID-Task-SA-Target-ExchangeOnline-MailboxGrantFullAccess
##############################################################
# Form mapping
$formObject = @{
    MailboxIdentity = $form.MailboxIdentity
    UsersToAdd      = $form.UsersToAdd
}

[bool] $IsConnected = $false

try {
    $null = Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)

    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop
    $IsConnected = $true

    foreach ($user in $formObject.UsersToAdd) {
        try {
            Write-Information "Executing ExchangeOnline action: [MailboxGrantFullAccess] for: [$($user.UserIdentity)] on mailbox: [$($formObject.MailboxIdentity)]"
            $splatParams = @{
                Identity        = $formObject.MailboxIdentity
                AccessRights    = 'FullAccess'
                InheritanceType = 'All'
                User            = $user.UserIdentity
                AutoMapping     = $true
            }
            $null = Add-MailboxPermission @splatParams -ErrorAction Stop
            $auditLog = @{
                Action            = 'UpdateResource'
                System            = 'ExchangeOnline'
                TargetIdentifier  = $($user.UserIdentity)
                TargetDisplayName = $($user.DisplayName)
                Message           = "ExchangeOnline action: [MailboxGrantFullAccess] for: [$($user.UserIdentity)] on mailbox: [$($formObject.MailboxIdentity)] executed successfully"
                IsError           = $false
            }
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information $auditLog.Message
        } catch {
            throw $_
        }
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $($user.UserIdentity)
        TargetDisplayName = $($user.DisplayName)
        Message           = "Could not execute ExchangeOnline action: [MailboxGrantFullAccess] for: [$($user.DisplayName)] on mailbox: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message), Details : $($_.Exception.ErrorDetails)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error $auditLog.Message
} finally {
    if ($IsConnected) {
        $exchangeSessionEnd = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
##############################################################
