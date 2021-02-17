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

REM This is a working example script to leverage headless Chrome CANARY Browser to to render HTML to PDF.
REM
REM Standard Chrome Browser always prints, usually undesired, header and footer and they can't be disabled.
REM The Chrome Canary supports an option print-to-pdf-no-header to disable printing of the file name and date
REM Enable this script in Mbox Viewer to disable undesired header and footer or 
REM different target directory is needed for PDF files.
REM NOTE: Chrome Canary comes from the developer node and may not be fully stable.
REM NOTE: From my expierence print-to-pdf option seems to be quite stable.
REM NOTE: Not clear whether in all cases HTML is fully render/loaded before printing to PDF

REM This script is invoked by Mbox Viewer and the full path to HTML file is passed as the first argument.
REM Configure path to this script via proper option in File -> Print Config.
REM To avoid suprises, script needs to be tested outside of the Mbox Viewer to make sure it works.

set FilePath=%1
set HTMLFilePath=%~1
Set HTMLdir=%~dp1
Set HTMLfile=%~nx1

REM The second parameter controls whether background color is removed or rather set to white
Set NoBackgroundColor=%2

REM Standard Chrome and Chrome Canary don't support no-backround option

if "%NoBackgroundColor%"=="--no-background" Set NoBackgroundColorOption="--default-background-color=0xffffff00"

REM Update path if needed
set ProgName=chrome.exe
set ProgDirectoryPath=C:\Users\%username%\AppData\Local\Google\Chrome SxS\Application
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

call "%CmdPath%"  --headless --disable-gpu --print-to-pdf-no-header --print-to-pdf="%PDFdir%\%HTMLNameBase%.pdf" "%HTMLFilePath%"

REM Replace "REM pause" with "pause" for testing.
REM pause
