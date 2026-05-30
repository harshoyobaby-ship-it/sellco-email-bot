@echo off
REM Quick send script — edit SUBJECT before running
set SUBJECT=Hello from SELLCO
python email_bot.py --subject "%SUBJECT%"
pause
