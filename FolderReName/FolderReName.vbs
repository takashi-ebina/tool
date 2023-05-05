'*********************************************************
' メイン処理
'*********************************************************
Option Explicit

Dim myDic: Set myDic = CreateObject("Scripting.Dictionary")
Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Dim arrTextLine

If Not WScript.Arguments.count = 1 Then
    WScript.Echo "引数の数が不正です"
    WScript.Quit()
End If

If Not objFileSys.FolderExists(WScript.Arguments(0)) Then
    WScript.Echo "フォルダではありません"
    WScript.Quit()
End If

Dim objReadStream: Set objReadStream = objFileSys.OpenTextFile(replace(Wscript.ScriptFullName,Wscript.ScriptName,"") & "FolderReName.txt", 1)
Do Until objReadStream.AtEndOfStream = True
    arrTextLine = Split(objReadStream.ReadLine, vbTab)
    myDic.add arrTextLine(0),arrTextLine(1)
LOOP
objReadStream.Close

call ReNameFolder(objFileSys.GetFolder(WScript.Arguments(0)))

set objFileSys = Nothing

msgbox "end"
'*********************************************************
'[概要]
' フォルダ名修正プロシージャ
'*********************************************************
sub ReNameFolder(folder)
    Dim f
    Dim element
    For Each f In folder.subFolders
        For Each element In myDic
            Dim objRe: Set objRe = New RegExp
            ' パターン
            objRe.Pattern = element
            ' 全体を検索
            objRe.Global = True
            ' 大文字小文字を区別しない
            objRe.IgnoreCase = True
            Dim strNewFolderName
	        strNewFolderName = objRe.Replace(f.Name, myDic.Item(element))
            If f.Name <> strNewFolderName Then
                f.Name = strNewFolderName
            End If
        Next
        ReNameFolder(objFileSys.GetFolder(f))
    Next

end sub
