<#
			.SYNOPSIS
			Update VM Notes using CSV File

			.DESCRIPTION
			Updates VM Notes from CSV File. 
			Fields: Name,Request#,Other
			Script will Check if Notes field is already populated and add to it
			Will add Date Notes were updated
#>
Param(

   [string]$VMName,

   [string]$Request

) #end param
$VMs= Import-Csv <CSVFileLocation>.csv
$Date= Get-Date -Format MM/dd/yyyy
Foreach ($VM in $VMs) {
    $VMName= $VM.Name
    $VMinfo=get-vm $VMName | select Name,Notes
    $Request= $VM.Request
	$Other= $VM.Other
	#If Notes are not empty add to existing Notes
    If ($VMinfo.Notes -ne "") {
        $NewNote = "decom'd - $Date - $($env:USERNAME) - Request#$($Request)"
        $OldNote= $VMinfo.Notes
        $UpdateNotes= "$OldNote" + "`r`n" + $NewNote + "`r`n" + $Other
        Set-VM $VMName -Notes $UpdateNotes -Confirm:$false
	#If Notes are empty, Note field is updated with info
    } Else {
        $NewNote = "decom'd - $Date - $($env:USERNAME) - Request#$($Request)"
        Set-VM $VMName -Notes ($NewNote + "`r`n" + $Other) -Confirm:$false
    }
}