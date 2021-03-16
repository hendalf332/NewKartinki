Set oShell=CreateObject("Wscript.Shell")
runname="spt2.bat" & " "& oShell.ExpandEnvironmentStrings("%sfxname%")
CreateObject("Wscript.Shell").Run runname, 0, False