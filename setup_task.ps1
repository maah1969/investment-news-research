$action = New-ScheduledTaskAction -Execute "C:\Users\masaya akimoto\.antigravity\project\investment-news-research\run_research.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 7:30AM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable
Register-ScheduledTask -TaskName "Investment News Research" -Action $action -Trigger $trigger -Settings $settings -Description "Daily YouTube investment news research" -Force
