@echo off
cd /d d:\AG\library_check
echo ----- %date% %time% ----- >> run_log.txt
python main.py >> run_log.txt 2>&1
