# HelloID-Task-SA-Target-ExchangeOnline-MailboxGrantFullAccess

## Prerequisites
Before using this snippet, verify you've met with the following requirements:
- [ ] The powershell EXO v3 module must be installed on the server running the Agent. See [Exchange-online-powershell-v2](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps) for instructions

- [ ] User defined variables: `ExchangeOnlineAdminUsername` and `$ExchangeOnlineAdminPassword` created in your HelloID portal.
  See also [Custom Variables](https://docs.helloid.com/en/variables/custom-variables.html)

## Description

This code snippet executes the following tasks:

1. Define a hash table `$formObject`. The keys of the hash table represent the properties of the room mailbox to be created, while the values represent the values entered in the form.

> To view an example of the form output, please refer to the JSON code pasted below.

```json
{
    "MailboxIdentity": "MySharedMailbox",
    "UsersToAdd": [
        {
            "UserIdentity": "84805b08-bc39-4fec-a79d-a4a6498dc7bz",
            "DisplayName": "JohnDoe"
        }
    ]
}
```

> :exclamation: It is important to note that the names of your form fields might differ. Ensure that the `$formObject` hashtable is appropriately adjusted to match your form fields.
> [See the Microsoft Docs page](https://learn.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps)

2. Constructs a powershell credential object from the supplied administrative username and password

3. Connects with the credentials to the Exchange online environment by means of the `Connect-ExchangeOnline` cmdlet

4. Calls the `Add-MailboxPermission` cmdlet to grant memberships.

5. Disconnects from the Exchange environment by means of the `Disconnect-ExchangeOnline` cmdlet
