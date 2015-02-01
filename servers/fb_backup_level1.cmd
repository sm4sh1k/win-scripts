@echo off

::Script for making incremental Firebird database backup (level 1) via nbackup utility,
:: archiving it with 7-zip and deleting old archives (35 days older).
::The full script set consist of three similar scripts for making incremental
:: everyday backups, weekly backups and full monthly backups.
::Author: Valentin Vakhrushev, 2014

SET nbackup="C:\Program Files (x86)\Firebird\Firebird_2_5\bin\nbackup.exe"
SET sevenzip="C:\Program Files\7-Zip\7z.exe"
SET database=X:\BASE\MAIN.FDB
SET archivepath=Y:\BACKUP\
SET dbuser=SYSDBA
SET dbpass=masterkey

SET year=%date:~-4%
SET month=%date:~3,2%
IF "%month:~0,1%" == " " SET month=0%month:~1,1%
SET day=%date:~0,2%
IF "%day:~0,1%" == " " SET day=0%day:~1,1%
SET archivename=MAIN_1_%year%-%month%-%day%

ECHO Creating database backup (level 1)...
%nbackup% -U %dbuser% -P %dbpass% -B 1 %database% %archivepath%%archivename%.nbk
IF %ERRORLEVEL% == 0 GOTO 7zip
EXIT 1

:7zip
ECHO Archiving backup file...
%sevenzip% a "%archivepath%%archivename%.7z" "%archivepath%%archivename%.nbk"
IF %ERRORLEVEL% == 0 GOTO finish
EXIT 1

:finish
ECHO Deleting original backup file...
DEL /f /q "%archivepath%%archivename%.nbk"
ECHO Deleting old archives (35 days older)...
CD /D "%archivepath%"
FORFILES /M MAIN_1_*-*-*.7z /D -35 /C "cmd /c if @isdir==FALSE del /f /q @file"
EXIT 0
