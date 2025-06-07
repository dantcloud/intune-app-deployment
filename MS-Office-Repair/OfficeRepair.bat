@echo off

:: Close all running Office applications
taskkill /f /im winword.exe
taskkill /f /im excel.exe
taskkill /f /im outlook.exe
taskkill /f /im powerpnt.exe
taskkill /f /im onenote.exe
taskkill /f /im mspub.exe

:: Wait a moment to ensure all Office apps are closed before starting the repair
timeout /t 5 /nobreak

:: Start Office repair and log output to a file
start /wait "" "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" scenario=Repair platform=x64 culture=en-gb RepairType=QuickRepair DisplayLevel=False >> C:\temp\office_repair_log.txt 2>&1

:: Add completion message to log
echo Office repair completed successfully on %date% at %time% >> C:\temp\office_repair_log.txt
