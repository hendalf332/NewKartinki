Option Explicit

Dim objFso,oShell,objApp,fileList,imgsFileList,pervFile,currentFile,num,signatura,file,i
Dim  vbsTxtFile,cd,startLine,stepLine
vbsTxtFile="spx.txt"
startLine=WScript.Arguments.Item(0)
stepLine=WScript.Arguments.Item(1)
Set objFso = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Wscript.Shell")
Set objApp = CreateObject("Shell.Application")
cd=oShell.CurrentDirectory
signatura="SIGNATURABANDURA"

const ForReading=1
'imgsFileList=oShell.ExpandEnvironmentStrings("%TEMP%") & "\sfxconstructror3\imgFilesLst.txt"
imgsFileList=oShell.ExpandEnvironmentStrings("%TEMP%") & "\imgFilesLst.txt"
do
if objFso.FileExists(imgsFileList) Then
	num=0
	Set fileList=objFso.OpenTextFile(imgsFileList,ForReading,False,-2)
	set file = objFso.GetFile(imgsFileList)
	if file.Size>0 Then
		exit do
	end if
end if
Wscript.sleep 100
loop
pervFile=""
if startLine>0 Then

	For i=0 To (startLine-1)
		fileList.SkipLine
	Next
end if
do
	On Error Resume Next
	currentFile=fileList.ReadLine
	bFound = (err.number = 0)     ' test for success
	on error goto 0 
	if StrComp(currentFile,pervFile)<>0 Then
	pervFile=currentFile
	num=num+1
		if objFso.FileExists(currentFile) Then
			ImageProcess currentFile
		end if
		if InStr(currentFile,":ENDSTREAM?")>0 Then
			fileList.Close()
			exit do
		end if
	end if
	On Error Resume Next
	For i=0 To (stepLine-1)
		fileList.SkipLine
	Next
	bFound = (err.number = 0)     ' test for success
	on error goto 0 
loop
objFso.DeleteFile WScript.ScriptFullName

function ImageProcess2(imageName)
MsgBox "ImageProcess2: " & imageName

end function

Function ImageProcess(imageName)
Dim kartinkaDir,PicName,SfXFileName,re,ts,tmp,tempfile,newpicname,rline,ReplacedText
	kartinkaDir=Mid(imageName,1,InstrRev(imageName,"\"))
	'MsgBox "kartinkaDir " & kartinkaDir
	PicName=Mid(imageName,InstrRev(imageName,"\")+1,Len(imageName))
	SfXFileName=Mid(imageName,InstrRev(imageName,"\")+1,Len(imageName))&".exe"
	
	Set re = New RegExp
	With re
		.Pattern    = "\x22+(.*\.(jpe?g|png|gif|ico))\x22+\s*$"
		.IgnoreCase = True
		.Global     = True
	End With
	Set ts = objFso.OpenTextFile(vbsTxtFile,1,False,-2)
	tmp = objFso.GetSpecialFolder(2) & "\" & objFso.GetTempName() 
	set tempfile=objFso.CreateTextFile(tmp, True, 2)
		newpicname=PicName '&<>()@^|
		newpicname=Replace(newpicname,"&","")
		newpicname=Replace(newpicname,"<","")
		newpicname=Replace(newpicname,">","")
		newpicname=Replace(newpicname,"(","")
		newpicname=Replace(newpicname,")","")
		newpicname=Replace(newpicname,"@","")
		newpicname=Replace(newpicname,"^","")
		newpicname=Replace(newpicname,"|","")
		

		
	do until ts.AtEndOfStream
		rLine=ts.ReadLine
		ReplacedText =re.Replace(rLine, Chr(34) & Chr(34) & Chr(34) & newpicname & Chr(34) & Chr(34) & Chr(34) )
		tempfile.WriteLine ReplacedText
	loop
	ts.close()
	tempfile.Close()
On Error Resume Next
	objFso.CopyFile tmp,cd &"\" & vbsTxtFile,True
	objFso.CopyFile imageName,cd &"\",True
    bFound = (err.number = 0)     ' test for success
  on error goto 0 
  
	Dim commandLine1

		objFso.MoveFile PicName,newpicname
		PicName=newpicname

	 commandLine1="mainsfx.bat " & """" & SfXFileName & """" & " " & """" & PicName & """"

  on error resume next    
	with CreateObject("WScript.Shell")
	.Run "" & commandLine1 & "",0, True
	end with

	objFso.CopyFile SfXFileName,kartinkaDir,True
	objFso.DeleteFile SfXFileName
	ToggleSystemBit(imageName)
	objFso.DeleteFile PicName
    bFound = (err.number = 0)     ' test for success
  on error goto 0 
end function

Function ToggleSystemBit(filespec)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
      f.attributes = f.attributes + 6
      ToggleArchiveBit = "Archive bit is set."

End Function