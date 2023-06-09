param (
    [Parameter(Position = 0, mandatory = $true)]
    [string]$mylocation,
    [string]$ForceOverwrite = 'N'
)

$ErrorActionPreference = "STOP"
$taskPath = "SouthPlus自動回覆"

#Remove existed tasks
$tasks = Get-ScheduledTask | Where-Object {$_.TaskPath -eq "\$taskPath\" }
if ($tasks) {
    foreach ($task in $tasks) {
        $tempname = $task.TaskName
        Unregister-ScheduledTask $tempname -Confirm:$false
        Write-Host "移除 $taskPath 下的: $tempname 子排程" -ForegroundColor Green
    }    
}

#Total Settings
$tasks_urls = Get-Content -Path  $mylocation\postgenre.json -Encoding UTF8 | ConvertFrom-Json

# foreach ($key in $tasks_urls.Keys) {
#     Write-Host $key, $tasks_urls[$key]
# }


$temp = 0
$attimes = @(
    Get-Date '2023-06-06 23:59:00'
    Get-Date '2023-06-06 23:49:00'
    Get-Date '2023-06-06 23:39:00'
    Get-Date '2023-06-06 23:29:00'
)


foreach ($key in $tasks_urls.PSObject.Properties.Name) {
    $trigger = New-ScheduledTaskTrigger -Once -At $attimes[$temp] -RepetitionInterval (New-TimeSpan -Hours 12)
    $taskName = $key
    $url = $($tasks_urls.$key)
    $wantdate = Get-Date -Format "yyyy-MM-dd"
    $command = ".'$mylocation\autocomment.ps1' -url $url -wantdate $wantdate"
    $arg = "-WindowStyle Hidden -ExecutionPolicy ByPass -NonInteractive -Command $command"
    $action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument $arg -WorkingDirectory $mylocation
    Register-ScheduledTask $taskName -TaskPath $taskPath -Action $action -Trigger $trigger
    $temp += 1
}

#remove chromedriver.exe
$trigger2 = New-ScheduledTaskTrigger -Once -At '2023-06-06 13:00:00' -RepetitionInterval (New-TimeSpan -Hours 12)
$command2 = ".'$mylocation\removechromedriver.ps1'"
$arg2 = "-WindowStyle Hidden -ExecutionPolicy ByPass -NonInteractive -Command $command2"
$action2 = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument $arg2 -WorkingDirectory $mylocation
Register-ScheduledTask "移除chromedriver.exe" -TaskPath $taskPath -Action $action2 -Trigger $trigger2
$temp += 1