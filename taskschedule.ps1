Param (
    [Parameter(Position=0,mandatory=$true)]
    [string]$mylocation,
    [string]$ForceOverwrite = 'N'
    )
$taskName = "SouthPlus自動回覆排程"
$taskPath = "SouthPlus自動回覆"
$ErrorActionPreference = "STOP"
$chkExist = Get-ScheduledTask | Where-Object { $_.TaskName -eq $taskName -and $_.TaskPath -eq "\$taskPath\" }
if ($chkExist) {
    if ($ForceOverwrite -eq 'Y' -or $(Read-Host "[$taskName] 已存在，是否刪除? (Y/N)").ToUpper() -eq 'Y') {
        Unregister-ScheduledTask $taskName -Confirm:$false 
    }
    else {
        Write-Host "放棄新增自動回覆排程" -ForegroundColor Red
        Exit 
    }
}

$mylocation = Get-Location
$url = ""
$wantdate = Get-Date -Format "yyyy-MM-dd"
$command = ".'$mylocation\autocomment.ps1' -url $url -wantdate $wantdate"
$arg = "-NoExit -ExecutionPolicy ByPass -NonInteractive -Command $command"
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument $arg -WorkingDirectory $mylocation

$trigger = New-ScheduledTaskTrigger -Once -At 2pm -RepetitionInterval (New-TimeSpan -Hours 2)
Register-ScheduledTask $taskName -TaskPath $taskPath -Action $action -Trigger $trigger