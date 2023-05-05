'*********************************************************
' メイン処理
'*********************************************************
Option Explicit

Const RESULT_DIR = ""
Const SEARCH_DIR = ""

Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Dim resultFile: Set resultFile = objFileSys.CreateTextFile(RESULT_DIR & "\" & "result.txt", true)

CALL SearchDir(SEARCH_DIR, resultFile)

msgbox "end"

'*********************************************************
'[概要]
' ファイルサイズ取得プロシージャー
'*********************************************************
sub SearchDir(branchPath, resultFile)

    resultFile.WriteLine "ファイルパス" & vbTab & "ファイル名" & vbTab & "ファイルサイズ" 

    For each f In objFileSys.GetFolder(branchPath).Files
        resultFile.WriteLine branchPath & vbTab & f.Name & vbTab & f.size
    Next
End sub