@echo off

@REM dos tips :)
@REM %~1         - expands %1 removing any surrounding quotes (")
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

@REM https://okular.kde.org/  PDF viewer capable of viewing very large PDF files


setlocal EnableExtensions  enabledelayedexpansion

if [%1]==[] goto :usage
if [%2]==[] goto :usage
if [%3]==[] ( set UserDataDir= ) else ( set UserDataDir=--user-data-dir="%~3")

set FirstMail=%~1
set LastMail=%~2

echo.
@echo FirstMail=%FirstMail% LastMail=%LastMail% UserDataDir=%UserDataDir%
echo.

set CmdPath=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe

set timeout=
set timeout=--timeout=1800000
@REM print-to-pdf exection time must be less the configured "timeout" otherwise some mails may not be printed to PDF
@REM Run type N.txt | findstr Print-to-pdf | sort  to check exection time

set user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
set user-agent=--user-agent="chrome/132.0.0.0"

@call :GetTimeInSeconds startTotalTime

set /a totalMails=%LastMail%-%FirstMail%+1

@REM #### Start converting range of HTML files fo PDF
for  /L %%i in (%FirstMail%,1,%LastMail%) do (

@set /a mailNumb=%%i-%FirstMail%+1
@set /a remainingFiles=%totalMails%-!mailNumb!

call :AddLeadingZeros %%i fileNumb 6
set fileName=m!fileNumb!

@echo File: !fileName!.htm  !DATE!  !TIME! file !mailNumb! of %totalMails% files remaining !remainingFiles!

@call :GetTimeInSeconds startTime

@del "%cd%\m%%i.pdf" >nul 2>&1

set html2pdf="%CmdPath%"  %UserDataDir% --headless=new %user-agent% %timeout% %virtual-time-budget% --disable-gpu --print-to-pdf-no-header --no-pdf-header-footer --print-to-pdf="%cd%\!fileName!.pdf" "%cd%\!fileName!.htm"
@echo !html2pdf!
!html2pdf!

@REM Helps to CTRL-C to break script and avoid looping over all steps
if NOT exist "%cd%\!fileName!.pdf" (
echo "%cd%\!fileName!.pdf" doean't exist
goto :eof
)


@call :GetTimeInSeconds endTime
@set /a executionTime=!endTime! - !startTime!
@call :ConvertSecondsToMinutesAndSeconds !executionTime!  minutes seconds

call :AddLeadingZeros !executionTime! deltaTime 10

@echo.
@echo !deltaTime! Print-to-pdf time for !fileName!.pdf: !executionTime! seconds or: !minutes! minutes and !seconds! seconds
@echo.
)

@REM #### End of convertion loop

@call :GetTimeInSeconds endTotalTime
@set /a executionTime=!endTotalTime! - !startTotalTime!
@call :ConvertSecondsToMinutesAndSeconds !executionTime!  min sec

@echo.
@echo Total time: !executionTime! seconds or: !min! minutes and !sec! seconds
@echo.

exit /b
goto :eof


@REM #####  Functions #####################

REM create and call myFunctions.bat instead of copy in all script files

:usage
@echo off
echo.
echo Usage: %0 firstMailIndex lastMailIndex [UserDataDirectory for Chrome/Edge instance when running multiple browser instances]
echo.
goto :eof
exit /b

:ConvertSecondsToMinutesAndSeconds
@echo off
SETLOCAL EnableExtensions enabledelayedexpansion
set totalSeconds=%~1
set /a minutes=%totalSeconds%/60
set /a minutesToSeconds=!minutes!*60
set /a seconds=!executionTime!-!minutesToSeconds!
( ENDLOCAL
set %~2=%minutes%
set %~3=%seconds%
)
goto :eof
exit /b

@REM works even if day changes within the same month
:GetTimeInSeconds
@echo off
SETLOCAL EnableExtensions enabledelayedexpansion

For /f "tokens=1-4 delims=/:." %%a in ("%TIME%") do (

	call :RemoveLeadingZeros %%a HH24
	call :RemoveLeadingZeros %%b MI
	call :RemoveLeadingZeros %%c SS
)

For /f "tokens=1-3 delims=/" %%a in ("%DATE%") do (
call :RemoveLeadingZeros %%b day
)

set /a timeInSeconds=!day!*24*3600 + %HH24%*3600 + %MI%*60 + %SS%

ENDLOCAL & set %~1=%timeInSeconds%
goto :eof
exit /b

:RemoveLeadingZeros
@echo off
SETLOCAL EnableExtensions enabledelayedexpansion
set s=%~1
:loop
if "!s:~0,1!" equ "0" (
	set s=!s:~1!
	if "!s!" neq "" (
		goto :loop
	)
)
if "!s!" equ "" (
	set s=0
)
ENDLOCAL & set %~2=%s%
goto :eof
exit /b

REM Add up to 12 zeros
:AddLeadingZeros
@echo off
SETLOCAL EnableExtensions enabledelayedexpansion
set num=000000000000%~1
set number=!num:~-%~3!
ENDLOCAL & set %~2=%number%
goto :eof
exit /b

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

:strlen
@echo off
set s=%1
setlocal enabledelayedexpansion
:strLen_Loop
  if not "!s:~%len%!"=="" set /A len+=1 & goto :strLen_Loop
(endlocal & set %2=%len%)
goto :eof
exit /b


