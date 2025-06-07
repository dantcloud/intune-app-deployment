@echo off

:: Check if the log file exists and delete it
if exist C:\temp\office_repair_log.txt (
    del /f /q C:\temp\office_repair_log.txt
    echo Log file deleted successfully.
) else (
    echo Log file not found.
)

:: Exit the script
exit /b
