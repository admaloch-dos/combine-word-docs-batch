@echo off
setlocal enabledelayedexpansion

::Intro text

echo.
echo   *************************************************
echo  ****** Florida Memory Combine Word Documents ******
echo   *************************************************
echo.

echo Instructions:
echo.
echo  --- 1. Locate the folder "add-word-docs-here"
echo  --- 2. Drag and drop the word documents you want to combine
echo  --- 3. (Optional) Add a name for the new file below
echo  --- 4. Click enter to merge the files
echo  --- 5. A new doc will show up in the current folder
echo.

:: Determine current location
set "main_location=%~dp0"

:: Set location to add word docs
set "word_doc_location=%main_location%\add-word-docs-here"

:: DocxMerge.exe file
set "merge_tool=%main_location%\DocxMerge.exe"

:: Set the output location
set "output_location=C:\Users\%USERNAME%\Desktop"


:: Prompt user for file name
set /p user_name_input="Enter a name for the new file or press enter: "

:: Set the file name with timestamp
set TIMESTAMP=%DATE:/=-%_%TIME::=-%
set TIMESTAMP=%TIMESTAMP: =%

IF "%user_name_input%"=="" (
    set "user_name_input=%USERNAME%_%TIMESTAMP%"
)

set "file_name=%user_name_input%"

:LOOP
:: Check if the word doc directory is empty
dir /b /s /a "%word_doc_location%\" | findstr .>nul || (
    echo Unable to locate any word documents.
    echo Add files to the, "Add Word Documents" folder, then double click enter to continue.
    pause >nul
    goto :LOOP
)

for /f "delims=" %%a in ('dir /b /a-d "%word_doc_location%\*"') do (
    set "file_type=%%~xa"
    goto :BREAK
)
:BREAK

set "files="

cd /d %word_doc_location%

for %%i in (*%file_type%) do set files=!files! "%%i"

"%merge_tool%" -i %files% -o "%main_location%\%file_name%.docx" -f
:: Delete contents of the word doc directory
: del /q "%word_doc_location%\*.*"
: rd /s /q "%word_doc_location%"


