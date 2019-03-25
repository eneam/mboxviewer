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
REM PDFBox doesn't seem to be able to merge very large number of PDF files and therefore the max number of PDF files to merge is configurable in this script.
REM Also, dos string seem to be limited to 8191 characters limiting the number of files to merge.
REM The PDF file name is derived from HTML file name and the length cen be limited by selecting File -> Print Config option in Mbox Viewer.
REM Script can be updated to leverage different and hopefully better PDF merge tools.

setlocal enabledelayedexpansion

set list=
set coptions=
set /A totalCount=0
set /A formattedTotalCount=0
set /A count=0
set /A maxPDFCount=50
set /A firstIndex=0
set /A listLength=0

set PDF_MERGE_DIR=PDF_MERGE

if NOT exist %PDF_MERGE_DIR% mkdir %PDF_MERGE_DIR%

for %%A in (*.pdf) do (

	if !count!==%firstIndex% set list="%%A"
	if NOT !count!==%firstIndex% set list=!list! "%%~nxA"

	set /A count=!count!+1
	set /A totalCount=!totalCount!+1

REM use temporarty file to determine string length, dos doesn't have a built in command
	echo.!list!>TempFileToDeterineStringLen
	for %%a in (TempFileToDeterineStringLen) do set /a TempFileToDeterineStringLen=%%~za -2


Set formattedTotalCount=0000000!totalCount!
Set formattedTotalCount=!formattedTotalCount:~-6%!

REM Duplicated code, no OR in DOS and using goto seem to break for loop
REM Max string length seem to be limited to 8191, 8192 - MAX_FILE_PATH ~ 7700
	set doPDF=n
	if !len! gtr 7700 (

		echo COUNT=!count!
		echo TOTAL_COUNT=!totalCount!
REM Replace next 3 lines with another HTML to PDF converion tool if desired
		set coptions=C:/Users/tata/Downloads/pdfbox-app-2.0.14.jar PDFMerger !list!  %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
		@echo java -jar !coptions!
		java -jar !coptions!

		set /A count=0
		set list=
		set doPDF=y
	)
	if NOT doPDF==y (
		if !count!==!maxPDFCount! (

			echo COUNT=!count!
			echo TOTAL_COUNT=!totalCount!
REM Replace next 3 lines with another HTML to PDF converion tool if desired
			set coptions=C:/Users/tata/Downloads/pdfbox-app-2.0.14.jar PDFMerger !list!  %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
			@echo java -jar !coptions!
			java -jar !coptions!

			set /A count=0
			set list=
		)
	)
)

if NOT %count%==0 (
	echo COUNT=!count!
	echo TOTAL_COUNT=%totalCount%
REM Replace next 3 lines with another HTML to PDF converion tool if desired
	set coptions=C:/Users/tata/Downloads/pdfbox-app-2.0.14.jar PDFMerger !list!  %PDF_MERGE_DIR%/all-!formattedTotalCount!.pdf
	@echo java -jar !coptions!
	java -jar !coptions!
)
if %count%==0 (
	echo COUNT=!count!
	echo TOTAL_COUNT=%totalCount%
	echo list !list!
)

if exist TempFileToDeterineStringLen DEL  TempFileToDeterineStringLen

pause