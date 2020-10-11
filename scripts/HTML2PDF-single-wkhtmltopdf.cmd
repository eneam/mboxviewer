@echo off

REM Few hints first :)
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

REM This is a working example script to leverage open source (LGPLv3) wkhtmltopdf tool to render HTML to PDF.
REM This script is the default configurable script in Mbox Viewer.
REM No need to update this script unless different target directory is needed for PDF files or 
REM other options such as header and footer must be configured for wkhtmltopdf tool.
REM NOTE: Using wkhtmltopdf, page header and footer can be controlled but not when using Chrome.
REM NOTE: Language and font support seem to be limited so make sure conversion works in your case.
REM NOTE: Not clear whether in all cases HTML is fully render/loaded before printing to PDF

REM This script is invoked by Mbox Viewer and the full path to HTML file is passed as the first argument.
REM Path to script can be configured and enabled by selecting proper option in File -> Print Config.
REM
REM To avoid suprises, script needs to be tested outside of the Mbox Viewer to make sure it works.
REM
REM wkhtmltopdf.exe version MUST be 0.12.6 or later  !!!!!!!!!!!!!!!!!
REM To verify the version, execute "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" -V


set FilePath=%1
set HTMLFilePath=%~1
Set HTMLdir=%~dp1
Set HTMLfile=%~nx1

REM The second parameter controls whether background color is removed
Set NoBackgroundColor=%2

if "%NoBackgroundColor%"=="--no-background" Set NoBackgroundColorOption=--no-background

REM Update path if needed
REM Download wkhtmltopdf from https://wkhtmltopdf.org/downloads.html
REM Usage link on how to control header and footer https://wkhtmltopdf.org/usage/wkhtmltopdf.txt
set ProgName=wkhtmltopdf.exe
set ProgDirectoryPath=C:\Program Files\wkhtmltopdf\bin
set CmdPath=%ProgDirectoryPath%\%ProgName%

REM Change the target directory for PDF files if needed.
set PDFdir=%HTMLdir%

REM echo Folder is: %HTMLdir%
REM echo Name is: %HTMLfile%
REM echo Path %HTMLFilePath%

For %%A in ("%HTMLfile%") do (
    Set HTMLNameBase=%%~nA
    Set HTMLNameExt=%%~xA
)
REM echo File Name Base is: %HTMLNameBase%
REM echo File Name Ext is: %HTMLNameExt%

del /Q "%PDFdir%\%HTMLNameBase%.pdf"

call "%CmdPath%" --log-level none --zoom 0.9 --enable-local-file-access %NoBackgroundColorOption% --footer-right "Page [page] of [toPage]" "%HTMLFilePath%" "%PDFdir%\%HTMLNameBase%.pdf"

REM Replace "REM pause" with "pause" for testing.
REM pause
