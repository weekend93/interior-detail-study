# 무재해 기록판 자동 갱신 - Windows 작업 스케줄러 등록 스크립트
# 사용법: 이 폴더에서 PowerShell을 열고  ->  .\register_task_scheduler.ps1
# (최초 1회만 실행하면 매일 지정한 시간에 자동으로 update_safety_board.py 가 실행됩니다.)

$ErrorActionPreference = "Stop"

$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    $pythonCmd = Get-Command py -ErrorAction SilentlyContinue
}
if (-not $pythonCmd) {
    Write-Host "Python을 찾을 수 없습니다. https://www.python.org 에서 설치 후 다시 실행해 주세요." -ForegroundColor Red
    exit 1
}

$scriptPath = Join-Path $PSScriptRoot "update_safety_board.py"
$action = New-ScheduledTaskAction -Execute $pythonCmd.Source -Argument "`"$scriptPath`"" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Daily -At 7:30am

Register-ScheduledTask `
    -TaskName "무재해기록판_자동갱신" `
    -Action $action `
    -Trigger $trigger `
    -Description "매일 아침 7:30에 무재해 기록판 PPTX의 달성일수/현재 날짜를 갱신하고 이미지를 바탕화면에 저장합니다." `
    -Force

Write-Host "등록 완료: 매일 오전 7:30에 자동 실행됩니다." -ForegroundColor Green
Write-Host "실행 시간을 바꾸려면 이 스크립트의 -At 7:30am 부분을 수정한 뒤 다시 실행하세요."
