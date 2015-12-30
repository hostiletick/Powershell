<#
			.SYNOPSIS 
			Creates a new User and mailbox for said User.

			.DESCRIPTION
			Verifies Tech Running script is an Admin and has Active Directory Module installed. 
			Creates a new User and mailbox for said User.
			Takes User input from questions asked, then acts accordingly.
			Creates User folder in location specified. 
			Asks if User needs Email Created and creates accordingly.

			.MODIFY
			Change all area that are labeled as <Insert Name>
			To Find all location find all instances of <Insert
#>
# Test Path for Loadouts and Results. Map Drive is needed.
If (!(Test-Path 'g:')){
	net use g: <Insert Script_Repository Path>
} Else {
	Write-Host "Drive is Mapped Already"
}
###################################################################
#	This is the AD User Creation section
###################################################################

# Ask for User input for Full Name, ID, Start Date, Title and Department
$uFullName= Read-Host "What is the new Users First and Last Name? Example: John Smith"
$uEmpID= Read-Host "What is the new Users Last Four of Social?  Example: 9999"
$uEmpHD= Read-Host "What is the new Users Hire Date. Example M/D/YYYY"
$uTitle= Read-Host "Enter Users Title"
$uDepartment= Read-Host "Enter Users Department"
	
# Asks to make sure this is not a Vendor or Service account that needs to be created in different OU.
$uADOU= Read-Host "What OU in AD do you want to put this new User in?  Default: --COURT USERS"
# Default OU is --COURT USERS Tech can press Enter to bybass typing anything to change OU
If ($uADOU -eq "") {
	$adLocation= "<Insert OU Name>"
	$uADOU= "<Insert FQDN of OU>"
} Else {
	# If Different OU is typed Script will Cancel and tell Tech to create manually.
	# Vendor and Service accounts do not get all the information that is needed for a Court User.
	write-host "This Account is not Normal Please Create Manually."
	Break
}

# create SamAccountName out of Full name First Int + LastName
$uFullName2= $uFullName.Split(' ')
$uFName= $uFullName2[0]
$uLName= $uFullName2[1]
$uName= $uFName.Substring(0,1).ToLower()
$uName+= $uLName.ToLower()

# check if SamAccountName is already used.
$uNameCheck= Get-ADUser -identity $uName
If ($uNameCheck.SamAccountName -eq $uName) {
    $tuName= Get-ADUser $uName
    Write-Host "The Username"$tuName.SamAccountName"already exist with"$tuName.Name"" -ForegroundColor Yellow -BackgroundColor Black
    $uName= Read-Host "Please type in the new SamAccountName to be used. Users Name is $uFullName"
} Else {
	# Script will continue if SamAccountName is not already used.
	$checked = "OK"
}

# Create User ID to place in Descrtiption
$uID= $uFname.Substring(0,1).ToLower()
$uID+= $uLname.Substring(0,1).ToLower()
$uID+= $uEmpID

# Create new User from input
$uHomeDrive= "<Insert Drive Letter:"
$uHomeDir= "<Insert Home DIR Locations"
$tempPass= "<Insert Password>"
 
###################################################
# Test Account Creation
# Uncomment this section below, if you need to just test the creation of variables or test updates on script #
################################################### 
##$ADCheck = Read-Host "Y/N"
## If Tech Answers Y script will continue to create Users AD Account
#If ($ADCheck -eq "Y") {
#	write-host "Creating Account"
#} Else {
	# If Tech Answers N Script will Terminated. 
#	write-host "Admin Terminated Script"
#	Break
#}
# End Test Account Creation #
###################################################

# Actually Creates
New-ADUser -Name "$uLName, $uFName" -AccountPassword (ConvertTo-SecureString -AsPlainText "$tempPass" -Force) `
-ChangePasswordAtLogon $true -Department "$uDepartment" -Description "ID=$uID, HD=$uEmpHD" `
-DisplayName "$uLName, $uFName" -Enabled $true -GivenName "$uFName" -HomeDirectory $uHomeDir `
-HomeDrive $uHomeDrive -Path "$uADOU" -SamAccountName $uName -Surname "$uLName" -Title "$uTitle" `
-UserPrincipalName ($uName + '@<Insert Domain Name>')

# add new User to default Groups
"<Insert Group Names devided by Comma>" | Add-ADGroupMember -Members $uName

###################################################
#	End User Creation Section
###################################################


###################################################
#	This Section Creates User Folder Under \\Frscrtfps\s$\2810\UserFolderss
#	Creates a folder with SamAccountName and under it a folder named private with only permissions to "Domain Admins and User"
###################################################

# Mount P Drive and test if already mapped.
#< Insert 
If (!(Test-Path p:)){ 
	net use p: <Instert User Folder Path>
} Else {
	Write-Host "Drive is Mapped Already"
}

# Change directory to make sure Folder is created under the P Drive
p:
cd \

# Creates Folder under UserFolders Under Main CourtHouse Folder on G Drive.
# Yes this is mapped to P Drive to insure correct location of the UserFolders location when running the script
$UserFolders = "P:\"
$uPrivatePath = "P:\$uName\"
$PrivatePath = "P:\$uName\$Dir"
$Dir = "Private"
	New-Item -Name $uName -ItemType Directory -Path $UserFolders
	New-Item -Name $Dir -ItemType Directory -Path $uPrivatePath

# Set Domain Admins Permissions on Private Folder
$DaUser = "Domain Admins"
$DaPerm = "FullControl"
$DaRule = "Allow"
$DaACL = Get-Acl "$PrivatePath"
$DaACL.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule `
("$DaUser", "$DaPerm", "ContainerInherit, ObjectInherit", "None", "$DaRule")))
	Set-Acl "$PrivatePath" $DaAcl

# Set Users Permissions on Private Folder 
$User = $uName
$Perm = "Modify"
$Rule = "Allow"
$ACL = Get-Acl "$PrivatePath"
$ACL.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule `
("$User", "$Perm", "ContainerInherit, ObjectInherit", "None", "$Rule")))
	Set-Acl "$PrivatePath" $Acl

# Removes Inherited Permissions from the folder.
$RIacl = Get-ACL -Path "$PrivatePath"
$RIacl.SetAccessRuleProtection($True, $False)
	Set-Acl -Path "$PrivatePath" -AclObject $RIacl

###################################################
#	End Folder Creation
###################################################

# Check if Use Needs Email Account
$EmailCheck = Read-Host "Does User Need Email - Y/N"
# If Tech Answers Y Script will continue and create user mailbox
If ($EmailCheck -eq 'Y') {
	write-host "Creating Exchange Mailbox"
} Else {
	"$uName`t$tempPass" | clip.exe
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("SamAccountName and Password copied to clipboard",0x2)
    Break
}

################################################################
#       This is the User Exchange Mailbox Creation section
################################################################

# Populate Mailbox Database and Set as "Users Group 8" to create the new Users email in
$Database= Get-MailboxDatabase -Identity "<Insert Database Name>"

# Create Users Mailbox with Alias to Match LastName, FirstName Look and in the correct Database
Enable-Mailbox -Identity "$adLocation/$uLName, $uFName" -Alias "$uName" -Database "$Database"

# Disable Mailbox Features
Set-CasMailbox -OWAEnabled $false -ActiveSyncEnabled $false -Identity "$adLocation/$uLName, $uFName"

# Check if Users Email has been created and Display it in Window.
$Email= Get-Mailbox -Identity "$uLName, $uFName"
	Write-Host "Email has been created"$Email.WindowsEmailAddress""


# Displayes New Created Users to SpreadSheet for sending to Management.
Get-ADUser $uName -Properties * | Select @{Name="Last, First";Expression={$_.Name}},@{Name="Email";Expression={$_.EmailAddress}}, `
    @{Name="UserName";Expression={$_.SamAccountName}},@{Name="Password";Expression={$tempPass}} | 
    Export-Csv -LiteralPath '<Insert Output Path and file name .csv>' -Append -NoTypeInformation -Force

# Copy Username and Password to clipboard
"$uName`t$tempPass" | clip.exe
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("SamAccountName and Password copied to clipboard",0x2)