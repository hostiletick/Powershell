<#
			.SYNOPSIS
			Allows Technician to Change or Delete a Current DHCP Reservation.

			.DESCRIPTION
			Collects Current IP Address and Location of Current IP Address
			Confirms that the IP Address is the Correct one then 
				asks the technician to Change or Delete it.
			Documents what was changed or deleted and then confirms

			.MODIFY
			Change all area that are labeled as <Insert Name>
			To Find all location find all instances of <Insert
			If you have more than two print locations and servers.
			Copy the ElseIf row and change the Location and Server Name
#>
# Sets Old Variables to $null so the logs are not enter with old data
$oIPadd = $null
$LOCINT = $null
$nMacadd = $null
$DateTime = $null

#Maps G Drive to document what is done with Reservation.
If (!(Test-Path g:)) {
	net use g: <Insert Script Repository Path>
} Else {
	Write-Host 'Drive is Mapped Already'
}
# Asks for Current IP Address you are looking to Modify.
$oIPadd = Read-Host 'Provide IP Address'

# Ask for Location of DHCP Reservation you are looking to
# modify and picks the correct server to connect to.
$LOCINT= Read-Host 'Provide Location <Insert Location Names devided by Comma>'
If ($LOCINT -eq '<Insert First Location from above>') {$LOCSERVER = '<Insert Server Name>'}
	ElseIf ($LOCINT -eq '<Insert Second Location from above>') {$LOCSERVER = '<Insert Server Name>'}
Else {Read-Host 'You Typed an unknown location please check again.'}

# Get Reservation information in preperation to change or delete.
$Reservation = Get-DhcpServerv4Reservation -ComputerName $LOCSERVER -IPAddress $oIPadd

# Check IP given on Server to see if a Reservation exists. If not it will ask technician for new IP.
If (!$Reservation) {
	$oIPadd = Read-Host 'IP Address not found please enter correct IP Address'
	Get-DhcpServerv4Reservation -ComputerName $LOCSERVER -IPAddress $oIPadd
} Else {
	$Reservation | FT -AutoSize
}
# Check Reservation to make technician has chosen the correct one.
$Continue = Read-Host 'Is The Above Reservation Correct Y/N'
If ($Continue -eq 'Y') {
	Write-Host 'Script will Continue'
} Else {
	Write-Warning 'Find Correct IPAddress and re-run this script.'
	Break
}
# Ask Technician to change or delete the Reservation.
# Documents what answer is given and resevation information to DHCP_Log.
$Tech= [Environment]::UserName
$DateTime= Get-Date -format 'MM-dd-yyyy HH:mm'
$QUEST= Read-Host 'Would you like to Change, or Delete the current Reservation C/D'
If ($QUEST -eq 'C') {
	$nMacadd = Read-Host 'Provide New Mac Address'
	$Change= 'Reservation Changed'
	# Documents that Reservation was changed in DHCP_Log. New Mac Address, Technician and Date included.
	ForEach-Object {
		Get-DhcpServerv4Reservation -ComputerName $LOCSERVER -IPAddress $oIPadd | Select @{N="Tech";E={$Tech}}, `
			@{N="Modified Date";E={$DateTime}},@{N="IPAddress";E={$_.IPAddress.IPAddressToString}}, `
			@{N="ScopeId";E={$_.IPAddress.IPAddressToString}},@{N="Clientid";E={$_.Clientid}},@{N="Name";E={$_.Name}}, `
			@{N="Description";E={$_.Description}},@{N="C/D";E={$Change}},@{N="New Mac";E={$nMacadd}} |
			Export-Csv "<Insert Full Output Path>" -NoTypeInformation -Append
	# Change DHCP Reservation
	Set-DhcpServerv4Reservation -IPAddress $oIPadd -ComputerName $LOCSERVER -ClientId $nMacadd
	}
	$cReservation = Get-DhcpServerv4Reservation -ComputerName $LOCSERVER -IPAddress $oIPadd
$cReservation | FT -AutoSize ; Write-Output 'Mac Address changed to above information!' -InformationAction Continue
} ElseIf ($QUEST -eq 'D') {
	$Sure = Read-Host 'Are you Sure you want to Delete the DHCP Reservation Y/N'
	If ($Sure -eq 'Y') {
		$Change= 'Reservation Deleted'
		# Documents that Reservation was deleted in DHCP_Log. Technician and Date are included.
		ForEach-Object {
			Get-DhcpServerv4Reservation -ComputerName $LOCSERVER -IPAddress $oIPadd | Select @{N="Tech";E={$Tech}}, `
				@{N="Modified Date";E={$DateTime}},@{N="IPAddress";E={$_.IPAddress.IPAddressToString}}, `
				@{N="ScopeId";E={$_.IPAddress.IPAddressToString}},@{N="Clientid";E={$_.Clientid}},@{N="Name";E={$_.Name}}, `
				@{N="Description";E={$_.Description}},@{N="C/D";E={$Change}},@{N="New Mac";E={$nMacadd}} |
				Export-Csv "<Insert Full Output Path>" -NoTypeInformation -Append
		# Delete DHCP Reservation
		Remove-DhcpServerv4Reservation -IPAddress $oIPadd -ComputerName $LOCSERVER
		}
	$Reservation | FT -AutoSize ; Write-Output 'Reservation Above Has Been Deleted!'
	}
} Else {
	Write-Host 'Script Terminated by Technician'
	Break
}