REM Initialize environment for x86
call "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvarsall.bat" x86
REM Embed manifest.xml in 7zSD.sfx
mt.exe -manifest manifest.xml -outputresource:"7zSD.sfx;#1"
REM Create the 7z file
7z.exe u Installer.7z "Application Files" SentFolderByFrom.vsto setup.exe Install.bat
REM Turn the 7z file into a self-extracting archive
REM copy /b 7zs.sfx + config.txt + Installer.7z SentFolderByFromInstaller.exe
copy /b 7zSD.sfx + config.txt + Installer.7z SentFolderByFromInstaller.exe