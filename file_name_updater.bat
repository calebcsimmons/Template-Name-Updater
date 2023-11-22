REM NAME: VCPTemplate File Name Updater
REM DESCRIPTION: Script to Update VCPTemplate Based On Currently Installed Software
REM AUTHOR: CALEB SIMMONS
REM DATE: 11/20/2023
REM Certain file names have been redacted (xxxx) due to privacy concerns

REM Possible Limitations:
REM 1.) If composites.txt file name is ever changed, relocated, or doesn't contain "xxxx" followed by version number
REM 2.) If this script is not placed within the 'commands' folder, it would need to be modified to accomodate its new location
REM 3.) If the VCPTemplate is ever renamed where it doesnt start with "VCPTemplate" or if its location is moved


@echo off
setlocal enabledelayedexpansion

rem Determine location of this current script
set "currentFolder=%~dp0"

rem Construct the file path by going up two levels and appending the relative path
rem Used relative instead of absolute due to variation in installation location / multiple software installs are present
set "compositesTxt=%currentFolder%..\..\composites.txt"

rem Check if the composites.txt file exists
if not exist "%compositesTxt%" (
    echo File not found: %compositesTxt%
    goto :EndScript
)

rem Read the content of the file and find the line containing "xxxxxxxx"
rem This line specifies the software version installed. Ex: "xxxx 9.4"
for /f "tokens=3*" %%a in ('type "%compositesTxt%" ^| find "xxxx"') do (
    set "versionNum=%%a"
)

rem Go to the folder that contains the VCPTemplate (...Composites x.x\xxx\xxx\xxx)
rem Used relative instead of absolute due to variation in installation location / multiple software installs are present

cd /d "%currentFolder%..\..\"
echo Current Directory before CD: %cd%
cd ".\xxx\xxx\xxx"
echo Current Directory after CD: %cd%

rem Find an Excel sheet starting with "VCPTemplate"
set "excelSheetCount=0"
for %%I in (VCPTemplate*) do (
    set /a "excelSheetCount+=1"
    set "excelSheet="%%I""

    rem Extract the file extension from the current template
    set "template_file_ext=%%~xI"
)

rem Error Handling in case no template, or multiple templates, are found
if %excelSheetCount% lss 1 (
    echo Error: No VCPTemplate found, cannot proceed.
    goto :EndScript
)

if %excelSheetCount% gtr 1 (
    echo Error: Multiple VCPTemplates found, cannot proceed.
    goto :EndScript
)

rem Create the updatedName by concatenating "VCPTemplate_" and versionNum
rem Used template_file_ext variable in case we ever use a different file ext for the template.
set "updatedName=VCPTemplate_!versionNum!!template_file_ext!"

rem Rename the Excel sheet to updatedName
ren !excelSheet! !updatedName!
echo Updated VCPTemplate Name: !updatedName!

:EndScript
pause
endlocal
