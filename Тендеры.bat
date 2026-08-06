@echo off
rem Запуск окна программы без консоли. cd — чтобы Python нашёл пакет tenders.
cd /d "%~dp0"
start "" pyw -3.12 -m tenders
