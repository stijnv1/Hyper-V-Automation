#
# CheckVMReplicaStatusAndMailCSV.ps1
#
param
(
    [Parameter(Mandatory=$True)]
    [string]$ClusterName,

    [Parameter(Mandatory=$True)]
    [switch]$ExportCSV,

    [Parameter(Mandatory=$False)]
    [string]$CSVPath,

    [Parameter(Mandatory=$True)]
    [string]$LogPath,

    [Parameter(Mandatory=$False)]
    [string]$SMTPServer,

    [Parameter(Mandatory=$False)]
    [string]$MailRecipient,

	[Parameter(Mandatory=$False)]
	[string]$MailSender

)

Function WriteToLog
{
	param
	(
		[string]$LogPath,
		[string]$TextValue,
		[bool]$WriteError
	)

	Try
	{
		#create log file name
		$thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
		$LogFileName = "ExportVMReplicaToCSV_$thisDate.log"

		#write content to log file
		if ($WriteError)
		{
			Add-Content -Value "[ERROR $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
		else
		{
			Add-Content -Value "[INFO $(Get-Date -DisplayHint Time)] $TextValue" -Path "$LogPath\$LogFileName"
		}
	}
	Catch
	{
	
	}

}

Try
{

    $nodes = Get-Cluster $ClusterName | Get-ClusterNode
    $VMReplicas = @()
    $VMReplicasMeasure = @()
    $VMReplicaOverview = @()

    foreach ($node in $nodes)
    {
        $VMReplicas += Get-VMReplication -ComputerName $node
        $VMReplicasMeasure += Measure-VMReplication -ComputerName $node
    }

    foreach ($VMReplica in $VMReplicas)
    {
        $VMObject = New-Object -TypeName PSObject
        $VMObject | Add-Member -MemberType NoteProperty -Name VMName -Value $VMReplica.Name
        $VMObject | Add-Member -MemberType NoteProperty -Name State -Value $VMReplica.State
        $VMObject | Add-Member -MemberType NoteProperty -Name Health -Value $VMReplica.Health
        $VMObject | Add-Member -MemberType NoteProperty -Name Frequency -Value $VMReplica.FrequencySec
        $VMObject | Add-Member -MemberType NoteProperty -Name PrimaryServer -Value $VMReplica.PrimaryServer
        $VMObject | Add-Member -MemberType NoteProperty -Name ReplicaServer -Value $VMReplica.ReplicaServer

        $VMReplicaMeasureObject = ($VMReplicasMeasure).Where{$_.Name -eq $VMReplica.Name}

        $VMObject | Add-Member -MemberType NoteProperty -Name LastReplicationTime -Value $VMReplicaMeasureObject.LReplTime
        $avgSize = $VMReplicaMeasureObject.AvgReplSize / 1MB
        $VMObject | Add-Member -MemberType NoteProperty -Name AverageReplicationSizeMB -Value $avgSize

        $VMReplicaOverview += $VMObject
    }

    if ($ExportCSV)
    {
        WriteToLog -LogPath $LogPath -TextValue "Creating CSV file for VM Replica Overview of Hyper-V Cluster $ClusterName ..."

        #create CSV file name
        $thisDate = (Get-Date -DisplayHint Date).ToLongDateString()
        $CSVFileName = "VMReplicaOverviewOfCluster_$ClusterName_$thisDate.csv"

        #export to CSV
        $VMReplicaOverview | Export-Csv -Path "$CSVPath\$CSVFileName" -NoTypeInformation

        #send mail with CSV as attachment
        Send-MailMessage -SmtpServer $SMTPServer -Attachments "$CSVPath\$CSVFileName" -From $MailSender -To $MailRecipient -Subject "VM Replica overview of cluster $ClusterName" -Body "Attached an overview of the VM Replica status in cluster $ClusterName"

        #delete CSV file
        Remove-Item -Path "$CSVPath\$CSVFileName"
    }
    else
    {
        $VMReplicaOverview | Out-GridView -Title "VM Replica Overview"
    }
}

Catch
{
    $ErrorMessage = $_.Exception.Message
	WriteToLog -LogPath $LogPath -TextValue "Error occured in WriteToLog function: $ErrorMessage" -WriteError $true
}