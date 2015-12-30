<#
			.SYNOPSIS 
			Grants FullAccess to users mailbox.
			Grants Send On Behalf Of to user if needed.

			.DESCRIPTION
			Grants FullAccess permissions to a mailbox chosen
			Asks admin for username and mailbox name.

			.MODIFY
			Change all area that are labeled as <Insert Name>
			To Find all location find all instances of <Insert
#>
#Cleans Variables setting them to $null.
$mbx= $null
$mb= $null
$gmb= $null
$uName= $null
$Check= $null
$dcheck= $null
$uCheck= $null

# Asks admin for Username and Mailbox Name.
# Mailbox Name can have wildcards added.
$uName= Read-Host "Enter Username of user getting access"

Do {
	$uCheck= Get-ADUser -filter * | where {$_.SamAccountName -like $uName}
	If ($uCheck -ne $null) {
		foreach ($User In $uCheck) {
			$dcheck= Read-Host ""$User.Name"<- is this User Correct Y/N"
			If ($dcheck -eq 'Y') {
				$uName= $User.SamAccountName
				Break
			} Else {
				Continue
			}
		}
	} Else {
		$uName= Read-Host "Username is not correct please re-enter"
	}
} Until ($dcheck -eq 'Y')


$gmb= Read-Host "Enter Mailbox Name. You Can use * to find correct spelling"

# Because Mailbox name excepts wildcards multple may be chosen.
# Below will get mailbox and then list each individually asking admin if that is correct.
# Script will not move on until question is answered Yes.
# Script will also make sure varialbe is not $Null and rerun question to get mailbox name needed.
Do {
	$mb= Get-Mailbox "$gmb"
	If ($mb -ne $null) {
		Foreach ($mbx In $mb) {
			$Check= Read-host "$mbx <- Is this the correct mailbox? Y/N"
			If ($Check -eq 'Y') {
				Break
			} Else {
				Continue
			}
		}
	} Else {
		$gmb= Read-Host "Mailbox is not found Enter Correct Spelling"
	}
} Until ($Check -eq 'Y')

# Checks if User already has access.
$ACheck= Get-MailboxPermission -Identity $mbx.Name | where {$_.User -like "<Insert Domain>\$uName"} | select User | ft -HideTableHeaders
If ($ACheck -eq $null) {
	Write-Host "Granting Access to"$mbx.Name""
} Else {
	Write-Host "User Already has Access"
	Break
}

# Adds user to mailbox with FullAccess permissions.
Add-MailboxPermission -Identity $mbx.Name -User $uName -AccessRights 'FullAccess'

# Asks admin if Send On Behalf Of access is needed. Grants depending on a yes no answer
$Send= Read-Host "Does User need access to Send On Behalf Of? Y/N"
If ($Send -eq 'Y') {
	set-Mailbox $mbx.Name -GrantSendOnBehalfTo @{add=”$uName”}
} Else {
	Write-Host "No extra permissions Granted"
}

# Asks Admin if you want to add another user to the same Mailbox.
Do {
	$uName= $null
	$uCheck= $null
	$dcheck= $null
	$Runagain= Read-Host "Do you need to add another user? Y/N"
	If ($Runagain -eq 'Y') {
		$uName= Read-Host "Enter Username of user getting access"
		Do {
			$uCheck= Get-ADUser -filter * | where {$_.SamAccountName -like $uName}
			If ($uCheck -ne $null) {
				Foreach ($User In $uCheck) {
					$dcheck= Read-Host ""$User.Name"<- is this User Correct Y/N"
					If ($dcheck -eq 'Y') {
						$uName= $User.SamAccountName
						Break
					} Else {
						Continue
					}
				}
			} Else {
				$uName= Read-Host "Username is not correct please re-enter"
			}
		} Until ($dcheck -eq 'Y')
		Add-MailboxPermission -Identity $mbx.Name -User $uName -AccessRights 'FullAccess'
		$Send= Read-Host "Does User need access to Send On Behalf Of? Y/N"
		If ($Send -eq 'Y') {
			set-Mailbox $mbx.Name -GrantSendOnBehalfTo @{add=”$uName”}
		} Else {
			Write-Host "No extra permissions Granted"
			Continue
		}
	} Else {
		Continue
	}
} Until ($Runagain -eq 'N')