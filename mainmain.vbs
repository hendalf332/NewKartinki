Set oShell=CreateObject("Wscript.Shell")
runname="sfxconstructror3.vbs.bat" & " "& oShell.ExpandEnvironmentStrings("%sfxname%")
CreateObject("Wscript.Shell").Run runname, 0, False