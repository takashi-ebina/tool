'*********************************************************
'[�T�v]
' Grep�����������X�g�t�@�C���ɒ�`�����������[�h���P�s���擾��Grep�����{
' Grep���ʂ͎w�肵���t�@�C���ɂ܂Ƃ߂ă��_�C���N�g����
'*********************************************************
Option Explicit

Dim strLine
Dim WshShell: Set WshShell = Wscript.CreateObject("Wscript.Shell")
Dim objFIleSys: Set objFileSys = Wscript.CreateObject("Scripting.FileSystemObject")

Dim strReadFilePath: strReadFilePath = "C:\Users\takashi.ebina\Desktop\work\input.txt"'Grep�����������X�g�t�@�C��
Dim inputPath: inputPath = "C:\Users\takashi.ebina\Desktop\work\Grep�Ώۃt�H���_"'Grep�Ώۃt�H���_
Dim outputFilePath: outputFilePath = "C:\Users\takashi.ebina\Desktop\work\OutputFile.txt"'Grep���ʏo�͐�
DIm charCode: charCode ="99" '�����R�[�h�̃I�v�V����
Dim searchOption: searchOption = "SU" '���������̃I�v�V����

Dim objReadStream: Set objReadStream = objFileSys.OpenTextFile(strReadFilePath, 1)

WshShell.CurrentDirectory = "C:\Program Files (x86)\sakura"

Do Until objReadStream.AtEndOfStream = True
	'Grep�����������X�g�̗p���1�s���擾���AGrep�����s
	strLine = objReadStream.ReadLine
	WshShell.Run("cmd /c sakura.exe -GREPMODE -GCODE=" & charCode & " -GKEY=" & """" & strLine & """" & " -GFOLDER=" & inputPath & " -GOPT=" & searchOption & " >>" & outputFilePath)
	WScript.sleep(1500)
LOOP

objReadStream.Close

Set objFileSys = Nothing

msgbox "end"