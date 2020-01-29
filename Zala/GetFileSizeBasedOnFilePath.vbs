'Pass file path inside " " so path use as single line.
' Ex: "C:\Folder_name\filename"

Dim Fso, Msg, FileObj, FilePath,C_Count
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.


' Store the arguments in a variable:
Set objArgs = Wscript.Arguments

'Count the arguments : For AA RPA input provide 2 more inline argument  , So minus those.
WScript.Echo objArgs.Count
C_Count=objArgs.Count - 2
WScript.Echo C_Count

'Read the inline file argument & create valid path
for i = 0 to C_Count step 1
  Wscript.Echo (i) & " :: " & objArgs.Item(i)
  If i <= 0 Then
	FilePath = FilePath & objArgs.Item(i)
  Else 
	FilePath = FilePath & " " & objArgs.Item(i)
  End If 
Next

Wscript.Echo FilePath 
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exits At The Specified Path
	Set FileObj = Fso.GetFile(FilePath) 'Returns "File" Object
	Msg = "File : " & FilePath & " Uses " & FileObj.Size & " Bytes" '.Size Property Returns Size Of The File In Bytes.
Else 'File Doesn't Exit.
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
' Use for output the file size as bytes.
WScript.StdOut.Write(Msg)
