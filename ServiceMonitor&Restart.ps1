#Script to monitor specific services on remote servers and restart them if found stopped. If restarted, email users to inform about this activity.

$Style = "<style>table, th, td {border: 1px solid;}</style>"
$Recipients = @("Example@email.com","Example2@email.com")

$ServerTable = [ordered]@{

	Server1 = "Service Name"
	Server2 = "Service Name"
	Server3 = "Service Name"
	Server4 = "Service Name"
	Server5 = "Service Name"
	Server6 = "Service Name"
	Server7 = "Service Name"
}
$ServiceInfo = @()

$ServerTable.GetEnumerator() | ForEach-Object {
    try {
        $Server = $_.key
        $Service = $_.value
        $Result = Get-Service -DisplayName $Service -ComputerName $Server -ErrorAction Stop
        if ($Result.Status -eq "Stopped") {
            $Parameters = @{
                ComputerName = $Server
                LogName = "System"
                Message = "The service entered the stopped state."
                Newest = 1
            }
            $ServiceStopTime = (Get-EventLog @Parameters).TimeGenerated
            try {
                Get-Service -DisplayName $Service -ComputerName $Server | Start-Service -ErrorAction Stop
                $Result = Get-Service -DisplayName $Service -ComputerName $Server
                $ServiceOutput = New-Object psobject
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Server Name" -Value $Server
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $Service
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Stop Time" -Value $ServiceStopTime
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Restart Time" -Value $(Get-Date)
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Status" -Value $Result.Status
                if (Test-Path C:\Temp\ServiceMonitor.csv){
                    $ServiceOutput | Export-Csv C:\Temp\ServiceMonitor.csv -NoTypeInformation -Append
                }
                else {
                    $ServiceOutput | Export-Csv C:\Temp\ServiceMonitor.csv -NoTypeInformation
                }
            }
            catch {
                $ServiceOutput = New-Object psobject
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Server Name" -Value $Server
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $Service
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Stop Time" -Value $ServiceStopTime
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Restart Time" -Value $(Get-Date)
                $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Status" -Value "Failed to Restart"
            }
            $ServiceInfo += $ServiceOutput
        }
    }
    catch {
        $ServiceOutput = New-Object psobject
        $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Server Name" -Value $Server
        $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $Service
        $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Stop Time" -Value "N/A"
        $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Service Restart Time" -Value "N/A"
        $ServiceOutput | Add-Member -MemberType NoteProperty -Name "Status" -Value "Cannot Connect to Server"
        $ServiceInfo += $ServiceOutput
    }
}

if ($ServiceInfo) {
    [String]$MessageBody = Import-Csv C:\Temp\ServiceMonitor.csv|ConvertTo-Html -Head $Style
    Send-MailMessage -to $Recipients `
                 -from no-reply@example.com `
                 -subject "SERVICE STATUS INFO: Service restarted" `
                 -BodyAsHtml `
                 -Body $MessageBody `
                 -SmtpServer xxx.xxx.xxx.xxx
}

if (Test-Path C:\Temp\NUIXServiceMonitor.csv) {
    Remove-Item -Path C:\temp\ServiceMonitor.csv -Force
}
