<#
			.SYNOPSIS
			Creates printer queue on remote print server

			.DESCRIPTION
			Creates a print queue on a remote print server
			Print server is chosen by printer location.

			.MODIFY
			Change all area that are labeled as <Insert Name>
			To Find all location find all instances of <Insert
			If you have more than two print locations and servers.
			Copy the ElseIf row and change the Location and Server Name
#>

# Asks admin what location the printer is being install and picks the appropriate server
$LOCINT= Read-Host 'Provide Location <Insert Location Names devided by Comma>'
If ($LOCINT -eq '<Insert First Location from above>') {$LOCSERVER = '<Insert Server Name>'}
	ElseIf ($LOCINT -eq '<Insert Second Location from above>') {$LOCSERVER = '<Insert Server Name>'}
Else {Read-Host 'You Typed an unknown location please check again.'}

# Aska aadmin what the printer name will be.
# This will be used to create queue and port
$prnName= Read-Host "Enter Printer Name. Do not add .court.fresno"

# Creates the printer port first to allow printer queue to choose it during creation process.
Add-PrinterPort -ComputerName $LOCSERVER -PortNumber 9100 -PrinterHostAddress ("$prnName" + "<Insert Domain Name>") -Name ("$prnName" + "<Insert Domain Name>")
Add-Printer -ComputerName $LOCSERVER -Name "$prnName" -DriverName "<Insert Print Driver Name>" -PortName ("$prnName" + "<Insert Domain Name>") -Shared -ShareName "$prnName"
