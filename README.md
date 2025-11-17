# Invoice Scraper - Automated Task Scheduler

A comprehensive guide for scheduling automated invoice scraping and processing tasks using Windows Task Scheduler via PowerShell.

## 📋 Overview

This project automates two daily tasks:
- **3:00 AM**: Run scraper only (`run_scraper_only.bat`)
- **7:00 AM**: Run full pipeline (`run_full_pipeline.bat`)

## 🚀 Quick Setup

### Prerequisites
- Windows OS with PowerShell
- Administrator privileges
- Batch files ready in your project directory

### Installation Steps

1. **Open PowerShell as Administrator**
   - Press `Win + X`
   - Select "Windows PowerShell (Admin)" or "Terminal (Admin)"

2. **Update File Paths**
   
   Replace `C:\Path\To\invoice-scraper\` with your actual project path in the commands below.

## 📅 Task Scheduling Commands

### ✅ Task 1: Scraper Only (3:00 AM Daily)

```powershell
$action = New-ScheduledTaskAction -Execute "C:\Path\To\invoice-scraper\run_scraper_only.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 3:00am
Register-ScheduledTask -TaskName "InvoiceScraper_3AM" -Action $action -Trigger $trigger -RunLevel Highest
```

### ✅ Task 2: Full Pipeline (7:00 AM Daily)

```powershell
$action = New-ScheduledTaskAction -Execute "C:\Path\To\invoice-scraper\run_full_pipeline.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00am
Register-ScheduledTask -TaskName "InvoiceFullPipeline_7AM" -Action $action -Trigger $trigger -RunLevel Highest
```

## 🛠️ Management Commands

### View All Scheduled Tasks

```powershell
Get-ScheduledTask
```

### View Specific Task Details

```powershell
Get-ScheduledTask -TaskName "InvoiceScraper_3AM"
```

### Run Task Manually

```powershell
Start-ScheduledTask -TaskName "InvoiceScraper_3AM"
```

```powershell
Start-ScheduledTask -TaskName "InvoiceFullPipeline_7AM"
```

### Delete a Task

```powershell
Unregister-ScheduledTask -TaskName "InvoiceScraper_3AM" -Confirm:$false
```

```powershell
Unregister-ScheduledTask -TaskName "InvoiceFullPipeline_7AM" -Confirm:$false
```

### Check Task Status

```powershell
Get-ScheduledTaskInfo -TaskName "InvoiceScraper_3AM"
```

## 📝 Example: Full Path Configuration

If your project is located at `C:\Projects\MaslahaScheduler\`, use:

```powershell
# Task 1
$action = New-ScheduledTaskAction -Execute "C:\Projects\MaslahaScheduler\run_scraper_only.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 3:00am
Register-ScheduledTask -TaskName "InvoiceScraper_3AM" -Action $action -Trigger $trigger -RunLevel Highest

# Task 2
$action = New-ScheduledTaskAction -Execute "C:\Projects\MaslahaScheduler\run_full_pipeline.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00am
Register-ScheduledTask -TaskName "InvoiceFullPipeline_7AM" -Action $action -Trigger $trigger -RunLevel Highest
```

## 🔧 Advanced Configuration

### Run with Specific User Credentials

```powershell
$action = New-ScheduledTaskAction -Execute "C:\Path\To\invoice-scraper\run_scraper_only.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 3:00am
$principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\Username" -LogonType Password -RunLevel Highest

Register-ScheduledTask -TaskName "InvoiceScraper_3AM" -Action $action -Trigger $trigger -Principal $principal
```

### Enable Logging

Add logging to your batch files:

```batch
@echo off
echo [%date% %time%] Starting scraper... >> C:\Path\To\logs\scraper.log
call your_script_here.bat >> C:\Path\To\logs\scraper.log 2>&1
echo [%date% %time%] Scraper finished >> C:\Path\To\logs\scraper.log
```

### Run on Multiple Days

```powershell
# Run only on weekdays
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 3:00am
```

### Add Multiple Triggers

```powershell
$trigger1 = New-ScheduledTaskTrigger -Daily -At 3:00am
$trigger2 = New-ScheduledTaskTrigger -Daily -At 9:00pm
Register-ScheduledTask -TaskName "InvoiceScraper_Multiple" -Action $action -Trigger $trigger1,$trigger2 -RunLevel Highest
```

## 📊 Monitoring

### View Last Run Result

```powershell
Get-ScheduledTaskInfo -TaskName "InvoiceScraper_3AM" | Select-Object LastRunTime, LastTaskResult
```

### Export Task Definition

```powershell
Export-ScheduledTask -TaskName "InvoiceScraper_3AM" | Out-File "C:\Backup\task_backup.xml"
```

### Import Task from Backup

```powershell
Register-ScheduledTask -Xml (Get-Content "C:\Backup\task_backup.xml" | Out-String) -TaskName "InvoiceScraper_3AM"
```

## ⚠️ Important Notes

1. **Administrator Rights**: Always run PowerShell as Administrator when creating tasks
2. **Path Format**: Use absolute paths (e.g., `C:\Projects\...`) not relative paths
3. **RunLevel Highest**: Ensures tasks run with elevated privileges
4. **Test Manually**: Always test your batch files manually before scheduling
5. **Check Logs**: Monitor task execution through Windows Event Viewer or custom logs

## 🐛 Troubleshooting

### Task Not Running?

1. **Check Task Status**:
   ```powershell
   Get-ScheduledTask -TaskName "InvoiceScraper_3AM" | Select-Object State
   ```

2. **Enable Task**:
   ```powershell
   Enable-ScheduledTask -TaskName "InvoiceScraper_3AM"
   ```

3. **Check Last Error**:
   ```powershell
   Get-ScheduledTaskInfo -TaskName "InvoiceScraper_3AM"
   ```

### Common Error Codes

- `0x0`: Success
- `0x1`: Incorrect function or path
- `0x41301`: Task is currently running
- `0x800710E0`: The operator or administrator has refused the request

## 📧 Email Notifications (Optional)

To add email notifications on failure, modify your batch file:

```batch
@echo off
your_script.bat
if %ERRORLEVEL% NEQ 0 (
    powershell -Command "Send-MailMessage -To 'admin@example.com' -From 'scheduler@example.com' -Subject 'Task Failed' -Body 'Invoice scraper failed' -SmtpServer 'smtp.gmail.com'"
)
```

## 📚 Resources

- [Microsoft Task Scheduler Documentation](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page)
- [PowerShell ScheduledTasks Module](https://docs.microsoft.com/en-us/powershell/module/scheduledtasks/)

## 📄 License

This documentation is provided as-is for educational and automation purposes.

---

**Author**: MaslahaScheduler Team  
**Last Updated**: November 2025
