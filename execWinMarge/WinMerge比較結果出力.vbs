'*********************************************************
'[概要]
' 指定したフォルダから、サブフォルダも含めてファイルを比較して、比較結果をhtm形式で出力します
' 同じフォルダ構成の同一名のファイルを対象とします
'
'[使い方]
' WinMerge比較結果出力.vbs 比較対象フォルダ1 比較対象フォルダ2 レポート出力先
'*********************************************************

Dim WshShell: Set WshShell=Wscript.CreateObject("Wscript.Shell")
Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Dim objParam: Set objParam = WScript.Arguments
Dim lngCnt: lngCnt = Len(objParam(0))

'WinMergeのインストールディレクトリに移動（PATHを通しているのであれば不要）
WshShell.CurrentDirectory = "C:\Program Files\WinMerge"

'WinMerge実行処理呼び出し
call execWinMerge(objParam(0), objParam(1), objParam(2))

msgbox "end"

'*********************************************************
'[概要]
' WinMerge実行プロシージャ
' 
'[引数]
' inputFilePath1：比較対象フォルダ1 
' inputFilePath2：比較対象フォルダ2 
' outputFilePath：レポート出力先
'*********************************************************
Sub execWinMerge(inputFilePath1,inputFilePath2, outputFilePath)

	'レポート出力先のフォルダ作成
	If not objFileSys.FolderExists(outputFilePath) Then
		objFileSys.CreateFolder(outputFilePath)
    End If
	
	'WinMerge実行
	For Each file In objFileSys.GetFolder(inputFilePath1).Files
		'比較対象ファイルの存在判定
		If objFileSys.FileExists(inputFilePath2 & "\" & file.Name) Then
			'同一ファイルが存在する場合にWinMerge実行
			inputFile1 =  inputFilePath1 & "\" & file.Name
			inputFile2 = inputFilePath2 & "\" & file.Name
			outputReport = outputFilePath & "\" & objFileSys.GetBaseName(file) & ".htm"
			WshShell.Run("WinMergeU.exe /e " & """" & inputFile1 & """" & " " & """" & inputFile2 & """" & " /minimize /noninteractive /u /or " & """" & outputReport & """")
		End If
	Next
	
	'サブフォルダに対してもWinMergeを実行
	For Each subfolder in objFileSys.GetFolder(inputFilePath1).Subfolders 
		If subfolder.Name <> "" Then
			'サブフォルダが存在する場合は、再帰処理を実行
			strSubfolderPath = Mid(subfolder, lngCnt + 2) 
			inputFilePath2 = objParam(1) & "\" & strSubfolderPath
			outputFilePath = objParam(2) & "\" & strSubfolderPath
			call execWinMerge(subfolder, inputFilePath2, outputFilePath)
		End If
	Next
End Sub
