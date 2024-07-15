@echo off

REM Путь к вашему виртуальному окружению
set VIRTUAL_ENV=.venv

REM Активируем виртуальное окружение
call %VIRTUAL_ENV%\Scripts\activate.bat

REM Запуск вашего Python файла
python __main__.py

REM Отключение виртуального окружения
deactivate

pause