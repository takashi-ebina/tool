'*********************************************************
'[概要]
' Grep検索条件リストファイルに定義した検索ワードを１行ずつ取得しGrepを実施
' Grep結果は指定したファイルにまとめてリダイレクトする
'*********************************************************
Option Explicit

Dim strLine
Dim WshShell: Set WshShell = Wscript.CreateObject("Wscript.Shell")
Dim objFIleSys: Set objFileSys = Wscript.CreateObject("Scripting.FileSystemObject")

Dim strReadFilePath: strReadFilePath = "C:\Users\takashi.ebina\Desktop\work\input.txt"'Grep検索条件リストファイル
Dim inputPath: inputPath = "C:\Users\takashi.ebina\Desktop\work\Grep対象フォルダ"'Grep対象フォルダ
Dim outputFilePath: outputFilePath = "C:\Users\takashi.ebina\Desktop\work\OutputFile.txt"'Grep結果出力先
DIm charCode: charCode ="99" '文字コードのオプション
Dim searchOption: searchOption = "SU" '検索条件のオプション

Dim objReadStream: Set objReadStream = objFileSys.OpenTextFile(strReadFilePath, 1)

WshShell.CurrentDirectory = "C:\Program Files (x86)\sakura"

Do Until objReadStream.AtEndOfStream = True
	'Grep検索条件リストの用語を1行ずつ取得し、Grepを実行
	strLine = objReadStream.ReadLine
	WshShell.Run("cmd /c sakura.exe -GREPMODE -GCODE=" & charCode & " -GKEY=" & """" & strLine & """" & " -GFOLDER=" & inputPath & " -GOPT=" & searchOption & " >>" & outputFilePath)
	WScript.sleep(1500)
LOOP

objReadStream.Close

Set objFileSys = Nothing

msgbox "end"