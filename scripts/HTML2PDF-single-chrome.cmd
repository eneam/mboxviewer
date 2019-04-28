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

REM This is a working example script to leverage headless Chrome Browser to to render HTML to PDF.
REM This is the same as hardcoded default processing in Mbox Viewer.
REM No need to configure and enable this script in Mbox Viewer unless new additional options are supported by the headless Chrome or 
REM different target directory is needed for PDF files.
REM NOTE: Not clear whether in all cases HTML is fully render/loaded before printing to PDF

REM This script is invoked by Mbox Viewer and the full path to HTML file is passed as the first argument.
REM Path to Chrome executable can be reconfigured if needed and enabled by selecting proper option in File -> Print Config.
REM To avoid suprises, script needs to be tested outside of the Mbox Viewer to make sure it works.

set FilePath=%1
set HTMLFilePath=%~1
Set HTMLdir=%~dp1
Set HTMLfile=%~nx1

REM Update path if needed
set ProgName=chrome.exe
set ProgDirectoryPath=C:\Program Files (x86)\Google\Chrome\Application
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

del "%PDFdir%\%HTMLNameBase%.pdf"

call "%CmdPath%"  --headless --disable-gpu --print-to-pdf="%PDFdir%\%HTMLNameBase%.pdf" "%HTMLFilePath%"

REM Replace "REM pause" with "pause" for testing.
REM pause
