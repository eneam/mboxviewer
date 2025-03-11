@echo off

REM dos tips :)
REM %~1         - expands %1 removing any surrounding quotes (")
REM %~f1        - expands %1 to a fully qualified path name
REM %~d1        - expands %1 to a drive letter only
REM %~p1        - expands %1 to a path only
REM %~n1        - expands %1 to a file name only
REM %~x1        - expands %1 to a file extension only
REM %~s1        - expanded path contains short names only
REM %~a1        - expands %1 to file attributes
REM %~t1        - expands %1 to date/time of file
REM %~z1        - expands %1 to size of file
REM The modifiers can be combined to get compound results:
REM %~dp1       - expands %1 to a drive letter and path only
REM %~nx1       - expands %1 to a file name and extension only

REM This script is working example script to merge multiple PDF files by leveraging open source Java Apache PDFBox tool.
REM PDFBox is licensed under Apache License version 2.0
REM DOS string seem to be limited to 8191 characters thus limiting the maximum number of PDF files to merge.
REM Script will merge files until DOS string limit or configurable max number of PDF files to merge is exceeded.
REM Work around is to merge PDF files multiple times, i.e. merge the merged files until only single merged PDF file remains.
REM Script can be updated to leverage different PDF merge tools.

setlocal enabledelayedexpansion

set list=
set coptions=
set /A totalCount=0
set /A formattedTotalCount=0
set /A count=0
set /A maxPDFCount=200
set /A firstIndex=0
set /A listLength=0

REM install java 8 or higher
REM Download java from  https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html
REM Update path to PDFBox jar file if needed
REM Download PDFBox from https://pdfbox.apache.org/
REM Command tools example https://pdfbox.apache.org/2.0/commandline.html

cd .
set PDFBox=pdfbox-app-3.0.4.jar
set PDFBoxDirectoryPath=%cd%
@REM Don't use full path, script breaks if path has UNICODE characters
@REM set PDFBoxJarFilePath=%PDFBoxDirectoryPath%/%PDFBox%
set PDFBoxJarFilePath=%PDFBox%

if NOT exist "%PDFBoxJarFilePath%" (

echo.
echo Invalid path to %PDFBoxJarFilePath% jar file.
echo Please install and/or update the PDFBox path in the batch file and re-run the script again.
echo.

pause
exit /b
)

set PDF_MERGE_DIR=PDF_MERGE

echo.
echo ################################################################
echo Current Directory=%PDFBoxDirectoryPath%
echo.

set javargs=
call :TotalPhysicalMemory sizeInGB 2>nul
if !sizeInGB! GEQ "16"  (

set /a stackSize=!sizeInGB!-8
@echo javargs=-XX:MaxMetaspaceSize=!stackSize!G -Xmx!stackSize!G -Xms!stackSize!G
echo.
set javargs=-XX:MaxMetaspaceSize=!stackSize!G -Xmx!stackSize!G -Xms!stackSize!G
)

if NOT exist %PDF_MERGE_DIR% mkdir %PDF_MERGE_DIR%

del "%PDF_MERGE_DIR%\*.pdf" 2>&1 2>nul

for %%A in ("*.pdf") do (

	if !count!==%firstIndex% set list=-i "%%A"
	if NOT !count!==%firstIndex% set list=!list! -i "%%~nxA"

	set /A count=!count!+1
	set /A totalCount=!totalCount!+1

REM use temporarty file to determine string length, dos doesn't have a built in command
	echo.!list!>TempFileToDeterineStringLen
	for %%a in (TempFileToDeterineStringLen) do set /a listLength=%%~za -2


Set formattedTotalCount=0000000!totalCount!
Set formattedTotalCount=!formattedTotalCount:~-6%!

REM Duplicated code, no OR in DOS and using goto seem to break for loop
REM Max string length seem to be limited to 8191, MAX_FILE_PATH ~ 260
	set doPDF=n
	if !listLength! gtr 7700 (
REM Replace next 3 lines with another HTML to PDF converion tool if desired
		set coptions=merge !list! %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
		@echo java %javargs% -jar "%PDFBoxJarFilePath%" !coptions!

		java %javargs% -jar  "%PDFBoxJarFilePath%" !coptions!

		set /A count=0
		set list=
		set doPDF=y
	)
	if NOT doPDF==y (
		if !count!==!maxPDFCount! (
REM Replace next 3 lines with another HTML to PDF converion tool if desired
			set coptions=merge !list! -o %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
			@echo java  %javargs% -jar "%PDFBoxJarFilePath%" !coptions!

			java %javargs% -jar "%PDFBoxJarFilePath%" !coptions!
			
			set /A count=0
			set list=
		)
	)
)

if NOT %count%==0 (
REM Replace next 3 lines with another HTML to PDF converion tool if desired
	set coptions=merge !list! -o %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
	@echo java  %javargs% -jar "%PDFBoxJarFilePath%" !coptions!
	
	java %javargs% -jar "%PDFBoxJarFilePath%" !coptions!
	
	pause
)
if %count%==0 (
	echo list !list!
)

if exist TempFileToDeterineStringLen DEL  TempFileToDeterineStringLen

echo.

pause
exit /b

REM create and call myFunctions.bat instead of copy in all script files

@REM in GBytes
:TotalPhysicalMemory
@echo off
SETLOCAL EnableExtensions  enabledelayedexpansion
for /F "tokens=*" %%A in ('systeminfo /fo list ^| findstr Total ^| findstr Physical ^| findstr Memory') do ( 
for /F "tokens=4 delims=, " %%i in ("%%A") do (
set size=%%i
)
)
ENDLOCAL & set %~1=%size%
goto :eof
exit /b