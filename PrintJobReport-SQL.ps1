<#
GeneratePrintJobAccountingReports.ps1
ver. 12-22-2015-01
This script reads event log ID 307 and ID 805 from the log "Applications and Services Logs > Microsoft > Windows > PrintService"
from the specified server and for the specified time period and then calculates print job and total page count data from these
event log entries.

It then writes the output to SQL Running Currently on CRT5744 
Requirements:
	- Ensure the .NET Framework 3.5 or later is installed:
	Add ".NET Framework 3.5.1" under the ".NET Framework 3.5.1 Features" option using the "Add Features" option of Server Manager, or
	- All Programs > Accessories > Windows PowerShell > right-click Windows PowerShell > Run as administrator...
		Import-Module ServerManager
		Add-WindowsFeature NET-Framework-Core
	- Enable and configure print job event logging on the desired print server:
	- start Devices and Printers > (highlight any printer) > Print server properties > Advanced
	- check "Show informational notifications for local printers"
	- check "Show informational notifications for network printers"
	- OK
	- start Event Viewer > Applications and Services Logs > Microsoft > Windows > PrintService
	- right-click Operational > Enable Log
	- right-click Operational > Properties > Maximum log size (KB): 65536 (was 1028 KB by default)
	- OK
	- Ensure that the user account used to run the script has write permission to the destination directory that will hold the
		output .CSV files ("D:\Scripts\" in the code below). Change the .CSV paths and filenames in the code below as desired.
	- If the print server is a remote server, ensure that the user account used to run the script has remote procedure call
		network access to the specified hostname, and that firewall rules permit such network access.
	- If the print server is logging events using a language other than English, customize the ID 805 message search string below
		to match the language-appropriate string used in the print server's event ID 805 event log message.
	- Ensure you have a SQL Database created. You can change perameter names or use the entries below to create tables
		- I have added the Table and View Creation SQL Scripts to my GitHub Repository Here. 
			
Usage:
	- see the PrintCommandLineUsage function, below 

Implementation notes:
	- Case of a HP LaserJet P2055dn printer using the HP Universal Printing PCL 5 (v5.2) driver
		The printer reports 0 copies on all jobs.
		If a print job reporting 0 copies is seen by the script, it will output a warning and then consider the affecte
			print job to be printed with 1 copy as a guess of what the actual number of copies was.
		The fix for this particular case was to upgrade the print driver to the HP Universal Printing PCL 5 (v5.5.0) driver.
	- Case of a HP LaserJet Pro 400 color printer model M451dn (CE957A) using the HP Universal Printing driver PCL 6 (v5.0.3),
		the HP Universal Printing PCL 5 (v5.2) driver and the HP Universal Printing PS driver (v.5.0.3):
		In all cases, this printer reports 1 copy of all jobs in Event ID 805, even when the user prints more than 1 copy of the job.
		There is no way for the script to detect this. It was discovered through observation.
		The fix was to clear the printer properties setting "Sharing > Render print jobs on client computers" (which is _enabled_ by default).
		With this change, the number of copies per job was reported accurately in the Windows event log.
		SUGGESTION: Check the generated .CSV file the first month and look check for printers that only ever report 1 copy of all jobs. These
			printers may need the work-around to render on the server-side.

History:
	- 2010-02 Original script written by Sh_Con at http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/007be664-1d8d-461c-9e0b-d8177106d4f8
	- 2011-05 Modified by BSOD2600 at http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/007be664-1d8d-461c-9e0b-d8177106d4f8
	- 2011-10 Modified by Tim Miller Dyck at PeaceWorks Technology Solutions to include the number of copies in page accounting by correlating with
		event ID 805, add target print server hostname and date parameters, add the by-user total pages report and switch encoding from
		Unicode to ASCII for better Excel .CSV compatibility.
		Thanks to Mennonite Central Committee Canada for sponsoring this additional development.
	- 2012-09 Modified by Tim Miller Dyck at PeaceWorks Technology Solutions to include a warning about print jobs reporting zero copies,
		add a warning about some print jobs incorrectly reporting one copy when more than one copy was printed, add the print job ID number
		to the .CSV output, and change commas in the print job name to underscores for more reliable .CSV parsing with some clients.
	- 2014-09 Modified by Tim Miller Dyck at PeaceWorks Technology Solutions to add additional warning logging and robustness for rare cases where
		event ID 805 messages are logged either 0 or more than 1 time for the same print job; add invalid document name character handling and PreviousDay
		improvements suggested by commentators at http://gallery.technet.microsoft.com/scriptcenter/Script-to-generate-print-84bdcf69/view/Discussions#content
	- 2015-12 Modified by Luke Ericksen to allow running on multiple Print Servers. Output is now to SQL Server for reporting. All relevant information is taken from
		event ID 307 and 805. Allowing Admin to report on all servers, users, total pages, and print color.
#>
Set-StrictMode -version 2
# Builds print server array to allow foreach to run through each. If more than one separate each by comma.
$prntservs = @("<Enter Server Name>")

# Get script start time to find System Event ID 104 to make sure log is cleared every time script runs.
$StartDate = Get-Date
$StartDate = $StartDate.ToString("yyyy-MM-dd HH:mm:ss")

# Runs through each print server to pull all printers
Foreach ($prntserv in $prntservs) {
	# the main print job entries are event ID 307 (use "-ErrorAction SilentlyContinue" to handle the case where no event log messages are found)
	$PrintEntries = Get-WinEvent -ErrorAction SilentlyContinue -ComputerName $prntserv -FilterHashTable @{ProviderName="Microsoft-Windows-PrintService"; ID=307}
	# the by-job number of copies are in event ID 805 (use "-ErrorAction SilentlyContinue" to handle the case where no event log messages are found)
	$PrintEntriesNumberofCopies = Get-WinEvent -ErrorAction SilentlyContinue -ComputerName $prntserv -FilterHashTable @{ProviderName="Microsoft-Windows-PrintService"; ID=805}

	# check for found data; if no event log ID 307 records were found, exit the script without creating an output file (this is not an error condition)

	#####
	# loop to parse ID 307 event log entries

	ForEach ($PrintEntry in $PrintEntries) {

		# get the date and time of the print job from the TimeCreated field
		$StartDate_Time = $PrintEntry.TimeCreated

		# convert the event log to an XML data structure
		#	Note that a print job document name that contains unusual characters that cannot be converted to XML will cause the .ToXml()
		#	method to fail so place a try/catch block around this code to address this condition. As an additional check, Windows Event Log Viewer
		#	will also fail to display the same event; the Details tab for the event will report "This event is not displayed correctly because the underlying XML is not well formed".
		#	Thanks to user Syncr0s for the report and fix posted at http://gallery.technet.microsoft.com/scriptcenter/Script-to-generate-print-84bdcf69/view/Discussions#content
		try {
			$entry = [xml]$PrintEntry.ToXml()
		}
		catch {
			# if ToXml has raised an error, log a warning to the console and the output file
			$Message = "WARNING: Event log ID 307 event at time $StartDate_Time has unparsable XML contents. This is usually caused by a print job document name that contains unusual characters that cannot be converted to XML. Please investigate further if possible. Skipping this print job entry entirely without counting its pages and continuing on..."
			$conn = New-Object System.Data.SqlClient.SqlConnection
			$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
			$conn.open()
			$cmd = New-Object System.Data.SqlClient.SqlCommand
			$cmd.connection = $conn
			$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[Warnings] (Warning,Server,PrinterName,Driver) VALUES('{0}','{1}','{2}','{3}')" -f $Message,$Server,$PrinterName,$PrintDriver
			$cmd.executenonquery()
			$conn.close()

			# and then immediately continue on with the next event ID 307 message, skipping the problem event log message
			Continue
		}

		# retreive the remaining fields from the event log UserData structure
		$307Time = $null
		$307EventID = $null
		$307ProcessID = $null
		$307ThreadID = $null
		$PrintJobId = $null
		$DocumentName = $null
		$UserName = $null
		$ClientPCName = $null
		$PrintColor = $null
		$PrintDriver = $null
		$PrinterName = $null
		$PrintSizeBytes =  $null
		$PrintPagesOneCopy = $null
	
		$307Time = $PrintEntry.TimeCreated
		$307EventID = $entry.Event.System.EventID
		$307ProcessID = $entry.Event.System.Execution.ProcessID
		$307ThreadID = $entry.Event.System.Execution.ThreadID
		$PrintJobId = $entry.Event.UserData.DocumentPrinted.Param1
		$DocumentName = $entry.Event.UserData.DocumentPrinted.Param2
		$UserName = $entry.Event.UserData.DocumentPrinted.Param3
		$ClientPCName = $entry.Event.UserData.DocumentPrinted.Param4
		$PrintDriver = $entry.Event.UserData.DocumentPrinted.Param5
		$PrinterName = $entry.Event.UserData.DocumentPrinted.Param6
		$PrintSizeBytes = $entry.Event.UserData.DocumentPrinted.Param7
		$PrintPagesOneCopy = $entry.Event.UserData.DocumentPrinted.Param8

		# get the user's full name from Active Directory
		if ($UserName -gt "") {
			$UserEntry= Get-ADUser $UserName
			$ADName = ($UserEntry.givenName + " " + $UserEntry.Surname)
		}
		# Write to SQL Server
		$conn = New-Object System.Data.SqlClient.SqlConnection
		$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
		$conn.open()
		$cmd = New-Object System.Data.SqlClient.SqlCommand
		$cmd.connection = $conn
		$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[307] (systemtime,eventid,processid,threadid,jobid,software,username,fullname,computername,printdriver,printername,printsize,numberofpages) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')" -f $307Time.ToString(),$307EventID,$307ProcessID,$307ThreadID,$PrintJobId,$DocumentName,$UserName,$ADName,$ClientPCName,$PrintDriver,$PrinterName,$PrintSizeBytes,$PrintPagesOneCopy
		$cmd.executenonquery()
		$conn.close()
		# get the print job number of copies by correlating with event ID 805 records
		#	the ID 805 record always is logged immediately before (that is, earlier in time) its related 307 record
		#	the print job ID number wraps after reaching 255, so we need to check both for a matching job ID
		#	and a very close logging time (within the previous 5 seconds) to its related event ID 307 record
		#	the print job ID match is based on a specific text string in the language of your Windows installation
		#	English search string: "Rendering job $PrintJobId."
		#	German search string: "Der Auftrag $PrintJobId wird gerendert." (thanks to user DJ83's post at
		#	http://gallery.technet.microsoft.com/scriptcenter/Script-to-generate-print-84bdcf69/view/Discussions#content)
		$PrintEntryNumberofCopies = $PrintEntriesNumberofCopies | Where-Object {$_.Message -eq "Rendering job $PrintJobId." -and $_.TimeCreated -le $StartDate_Time -and $_.TimeCreated -ge ($StartDate_Time - (New-Timespan -second 5))}

		# check for the expected case of exactly one matching event ID 805 event log record for the source event ID 307 record
		#   if this is true then extract the number of print job copies for the matching print job
		if (($PrintEntryNumberofCopies | Measure-Object).Count -eq 1) {
			# retrieve the remaining fields from the event log contents
			$entry2 = [xml]$PrintEntryNumberofCopies.ToXml()
			$805Time = $null
			$805EventID = $null
			$805ProcessID = $null
			$805ThreadID = $null
			$Server = $null
			$UserID = $null
			$805JobID = $null
			$805Jobsize = $null
			$icmmethod = $null
			$PrintColor = $null
			$XRes = $null
			$YRes = $null
			$Quality = $null
			$NumberOfCopies = $null
			$TTOption = $null

			$805Time = $PrintEntryNumberofCopies.TimeCreated
			$805EventID = $entry2.Event.System.EventID
			$805ProcessID = $entry2.Event.System.Execution.ProcessID
			$805ThreadID = $entry2.Event.System.Execution.ThreadID
			$Server = $entry2.Event.System.Computer
			$UserID = $entry2.Event.System.Security.UserID
			$805JobID = $entry2.Event.UserData.RenderJobDiag.JobId
			$805Jobsize = $entry2.Event.UserData.RenderJobDiag.GdiJobSize
			$icmmethod = $entry2.Event.UserData.RenderJobDiag.ICMMethod
			$PrintColor = $entry2.Event.UserData.RenderJobDiag.Color
			$XRes = $entry2.Event.UserData.RenderJobDiag.XRes
			$YRes = $entry2.Event.UserData.RenderJobDiag.YRes
			$Quality = $entry2.Event.UserData.RenderJobDiag.Quality
			$NumberOfCopies = $entry2.Event.UserData.RenderJobDiag.Copies
			$TTOption = $entry2.Event.UserData.RenderJobDiag.TTOption

			If ($PrintColor -gt '1') {
				$PrintColor = "Color"
			} Else {
				$PrintColor = "Black"
			}
			# some flawed printer drivers always report 0 copies for every print job; output a warning so this can be investigated
			#	further and set copies to be 1 in this case as a guess of what the actual number of copies was
			if ($NumberOfCopies -eq 0) {
				$NumberOfCopies = 1
				$Message = "WARNING: Printer $PrinterName recorded that print job ID $PrintJobId was printed with 0 copies. This is probably a bug in the print driver, change the print driver"
				$conn = New-Object System.Data.SqlClient.SqlConnection
				$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
				$conn.open()
				$cmd = New-Object System.Data.SqlClient.SqlCommand
				$cmd.connection = $conn
				$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[Warnings] (Warning,Server,PrinterName,Driver) VALUES('{0}','{1}','{2}','{3}')" -f $Message,$Server,$PrinterName,$PrintDriver
				$cmd.executenonquery()
				$conn.close()
			}
		}
		# otherwise, either no or more than 1 matching event log ID 805 record was found
		#   both cases are unusual error conditions so report the error but continue on, assuming one copy was printed
		else {
			$NumberOfCopies = 1
			$Message = "WARNING: Printer $PrinterName recorded that print job ID $PrintJobId had $(($PrintEntryNumberofCopies | Measure-Object).Count) matching event ID 805 entries in the search time range from $(($StartDate_Time - (New-Timespan -second 5))) to $StartDate_Time. Logging this as a warning as only a single matching event log ID 805 record should be present. Please investigate further if possible. Guessing that 1 copy of the job was printed and continuing on..."
			$conn = New-Object System.Data.SqlClient.SqlConnection
			$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
			$conn.open()
			$cmd = New-Object System.Data.SqlClient.SqlCommand
			$cmd.connection = $conn
			$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[Warnings] (Warning,Server,PrinterName,Driver) VALUES('{0}','{1}','{2}','{3}')" -f $Message,$Server,$PrinterName,$PrintDriver
			$cmd.executenonquery()
			$conn.close()
		}

		# calculate the total number of pages for the whole print job
		$TotalPages = [int]$PrintPagesOneCopy * [int]$NumberOfCopies
		
		# Write to SQL
		$conn = New-Object System.Data.SqlClient.SqlConnection
		$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
		$conn.open()
		$cmd = New-Object System.Data.SqlClient.SqlCommand
		$cmd.connection = $conn
		$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[805] (systemtime,eventid,processid,threadid,server,userid,jobid,jobsize,icmmethod,printcolor,xres,yres,quality,copies,ttoption,totalpages) VALUES(@time,@id,@procid,@threadid,@server,@userid,@jobid,@size,@method,@color,@xres,@yres,@quality,@ncopies,@option,@tpapes)";
		$cmd.Parameters.Add("@time", $805Time.ToString());
		$cmd.Parameters.Add("@id", $805EventID);
		$cmd.Parameters.Add("@procid", $805ProcessID);
		$cmd.Parameters.Add("@threadid", $805ThreadID);
		$cmd.Parameters.Add("@server", $Server);
		$cmd.Parameters.Add("@userid", $UserID);
		$cmd.Parameters.Add("@jobid", $805JobID);
		$cmd.Parameters.Add("@size", $805Jobsize);
		$cmd.Parameters.Add("@method", $icmmethod);
		$cmd.Parameters.Add("@color", $PrintColor);
		$cmd.Parameters.Add("@xres", $XRes);
		$cmd.Parameters.Add("@yres", $YRes);
		$cmd.Parameters.Add("@quality", $Quality);
		$cmd.Parameters.Add("@ncopies", $NumberOfCopies);
		$cmd.Parameters.Add("@option", $TTOption);
		$cmd.Parameters.Add("@tpapes", $Totalpages);
		$cmd.ExecuteNonQuery();
		$conn.close()
	}
	wevtutil cl Microsoft-Windows-PrintService/Operational /r:$prntserv
	# Get script end time to find System Event ID 104 to make sure log is cleared every time script runs.
	# If the $EndDate Time is the same as the Log Entry then you will not write the entry in the SQL database because
	# it does not see the entry when searching for it with the same time stamp. 
	# Added a Time Span of 30 seconds for when script runs so fast the $EndDate time is the same as the log entry. 
	$ts = New-TimeSpan -Seconds 30
	Get-Date
	$EndDate = (Get-Date) + $ts
	$EndDate = $EndDate.ToString("yyyy-MM-dd HH:mm:ss")
	$EndDate
	# Get Event ID 104 and Write to SQL Server for logging each Print Server.
	$clearlog = Get-WinEvent -ErrorAction SilentlyContinue -ComputerName $prntserv -FilterHashTable @{ProviderName="Microsoft-Windows-Eventlog"; StartTime=$StartDate; EndTime=$EndDate; ID=104}
	$104UserID = $clearlog.UserId
	# I have changed the SID of the users that runs all my scripts to the acually AD Name so I know who is running it. 
	If ($104UserID -eq '<Enter SID>') {
		$104UserID = '<Enter UserName>'
	} Else {
		$104UserID = Write-Host "$104UserID not srunner"
	}
	$clearlogtime = $clearlog.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
	$104ID = $clearlog.Id
	$TaskName = $clearlog.TaskDisplayName
	$104Server = $clearlog.MachineName
	# Write Clear Log Event to database for each Server.
	$conn = New-Object System.Data.SqlClient.SqlConnection
	$conn.ConnectionString = "Data Source=<Enter SQL Server>;Initial Catalog=PrintReport;Integrated Security=SSPI;"
	$conn.open()
	$cmd = New-Object System.Data.SqlClient.SqlCommand
	$cmd.connection = $conn
	$cmd.commandtext = "INSERT INTO [PrintReport].[dbo].[104] (systemtime,eventid,taskname,server,userid) VALUES(@time,@id,@taskname,@server,@userid)";
	$cmd.Parameters.Add("@time", $clearlogtime);
	$cmd.Parameters.Add("@id", $104ID);
	$cmd.Parameters.Add("@taskname", $TaskName);
	$cmd.Parameters.Add("@server", $104Server);
	$cmd.Parameters.Add("@userid", $104UserID);
	$cmd.ExecuteNonQuery();
	$conn.close()
	$cmd.ExecuteNonQuery();
	$conn.close()
}