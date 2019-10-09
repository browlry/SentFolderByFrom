REM Create the folder, if it does not exist.
if not exist "%appdata%\browlry\SentFolderByFrom" mkdir "%appdata%\browlry\SentFolderByFrom"
REM Move all these files and folders to the appdata folder
robocopy . "%APPDATA%\browlry\SentFolderByFrom" /s
REM Launch setup.exe
start "" "%APPDATA%\browlry\SentFolderByFrom\setup.exe"