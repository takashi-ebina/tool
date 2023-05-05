'*********************************************************
' ÉÅÉCÉìèàóù
'*********************************************************
Option Explicit

Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Dim strBaseDirPath: strBaseDirPath = objFileSys.GetAbsolutePathName(Wscript.Arguments(0))
Dim objDir: Set objDir = objFileSys.getFolder(strBaseDirPath)

For Each file In objDir.Files
    Dim objInput: Set objInput = CreateObject("ADODB.Stream")
    Dim objOutput: Set objOutput = CreateObject("ADODB.Stream")

    objInput.Charset = "Shift_JIS"
    objInput.Open
    objInput.LoadFormFile file

    tmpStr = objInput.ReadText

    objOutput.Charset = "Shift_JIS"
    objOutput.Open
    objOutput.LoadFormFile tmpStr

    objInput.Close

    objOutput.saveToFile file, 2

    objOutput.Close
Next
msgbox "end"

