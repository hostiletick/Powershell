
$ModCheckAD= Get-Module ActiveDirectory
If ($ModCheckAD -eq $null) {
    # Import ActiveDirectory Module
    Import-Module ActiveDirectory
}
$ModCheckVM= Get-Module VMware*
If ($ModCheckVM -eq $Null) {
    # Import VMware Module
    Get-Module VMware* -ListAvailable | Import-Module
}
$VICheck= $global:DefaultVIServer[0].Name
If ($VICheck -ne '<Enter Vcenter>') {
    Connect-VIServer <Enter Vcenter>
}
Remove-Item $env:HOMEPATH\Documents\results\VMinfo.csv -Force
#Collect Server name
#$SNames= Read-Host 'Enter Server Name: '
#$userinput = Read-Host "Enter List Comma Seperated"
#$SNames= $userinput.Split(",").Trim(" ")
$SNames= Get-Content \\tsclient\y\imports\Serverlist.txt
$Path= "C:\Users\LUKE064-DSA\Documents\results\ServerInfo.txt"
Foreach ($SName in $SNames) {
    # Collect VM Info          
    $nwINFO = Get-WmiObject -ComputerName $SName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null } `
        #| Select-Object DNSHostName,Description,IPAddress,IpSubnet,DefaultIPGateway,MACAddress,DNSServerSearchOrder | format-Table * -AutoSize  
        #| Select-Object DNSHostName,Description,IPAddress,IpSubnet,DefaultIPGateway,MACAddress,DNSServerSearchOrder 
    $nwServerName = $nwINFO.DNSHostName 
    $nwDescrip = $nwINFO.Description 
    $nwIPADDR = $nwINFO.IPAddress 
    $nwSUBNET = $nwINFO.IpSubnet 
    $nwGateWay = $nwINFO.DefaultIPGateway 
    $nwMacADD = $nwINFO.MACAddress 
    $nwDNS = $nwINFO.DNSServerSearchOrder
    $CPU= Get-WmiObject -ComputerName $Sname win32_processor 
    $VMCPU= $CPU.Count
    $OSInfo= Get-WmiObject -ComputerName $Sname -class Win32_OperatingSystem | Select @{N="Mem";E={$_.TotalVisibleMemorySize / 1MB}}, Caption
    $VMMEM= $OSInfo.Mem
    $VMOSName= $OSInfo.Caption
    $VMVLAN= (Get-VDPortGroup -VM $SName -ErrorAction SilentlyContinue | select Name).Name
    $VMAR="Fill"
    $VCILocal="Fill"
    $VMCALocal="Fill"
    $Test= "Fun"
    $VMHDCap= ((Get-HardDisk -VM $SName | Measure-Object -Sum CapacityGB).Sum)
    $VMinfo= Get-VM $SName -ErrorAction SilentlyContinue | Select Name,Guest,MemoryGB,NumCpu,Powerstate,@{N="Cluster";E={Get-Cluster -VM $_ -ErrorAction SilentlyContinue}},@{N="VMhost";E={Get-VMHost -VM $_ -ErrorAction SilentlyContinue}}, `
        @{N="Datastore";E={Get-Datastore -VM $_ -ErrorAction SilentlyContinue}}
    $VMDatastore= $VMinfo.Datastore.Name
    $VMClusterName= ($VMinfo.VMhost,$VMinfo.Cluster -join"\")
    $nwDNSIPs= ($nwDNS -join"/")
    $after="CopyAfter"
    $Power=$VMinfo.PowerState
    $InitialCNtest = (Get-ADComputer $SName -Properties CanonicalName).CanonicalName -Split ("/")
    If ($InitialCNtest.CanonicalName -ne $Null) {
        $InitialCN= $InitialCNtest
    } Else {
        $InitialCN= (Get-ADComputer $SName -Server slcdc01.corp.questar.com -Properties CanonicalName).CanonicalName -Split ("/")
    }
    $ParentOU = $InitialCN[0..$($InitialCN.Count-2)] -Join "/"
    $Test2= $Test | Select @{N="ServerName";E={$SName}},@{N="Description";E={$nwDescrip}},@{N="IP";E={$nwIPADDR}},@{N="Subnet";E={$nwSUBNET}}, `
        @{N="Gateway";E={$nwGateWay}},@{N="DNS";E={$nwDNS}},@{N="after";E={$after}},@{N="CIP";E={$VMinfo.Guest.IPAddress}},@{N="CSubnet";E={$nwSUBNET}}, `
        @{N="CGateway";E={$nwGateWay}},@{N="DNSIPs";E={$nwDNSIPs}},@{N="OSName";E={$VMinfo.Guest.OSFullName}},@{N="Mem";E={$VMinfo.MemoryGB}},@{N="CPU";E={$VMinfo.NumCpu}}, `
        @{N="HDCAP";E={$VMHDCap}},@{N="OU";E={$ParentOU}},@{N="VMAR";E={$VMAR}},@{N="VCILocal";E={$VCILocal}},@{N="Cluser";E={$VMClusterName}}, `
        @{N="VMDatastore";E={$VMDatastore}},@{N="VMOSName";E={$VMOSName}},@{N="VMVLAN";E={$VMVLAN}},@{N="VMHW";E={$VMHW}}, `
        @{N="VmTools";E={$VMToolsVer}},@{N="Mac";E={$nwMacADD}},@{N="Power";E={$Power}} `
        | Export-Csv $env:HOMEPATH\Documents\results\VMinfo.csv -Append -NoTypeInformation
}
# Copy new results to Y Drive with Date stamp.
$date=Get-Date -Format MM-dd-yy
$newname= "vminfo-" + $date + ".csv"
$cPath= "<PATH>" + $newname
Copy <csvPath> $cPath