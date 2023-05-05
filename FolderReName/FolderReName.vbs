'*********************************************************
' ���C������
'*********************************************************
Option Explicit

Dim myDic: Set myDic = CreateObject("Scripting.Dictionary")
Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Dim arrTextLine

If Not WScript.Arguments.count = 1 Then
    WScript.Echo "�����̐����s���ł�"
    WScript.Quit()
End If

If Not objFileSys.FolderExists(WScript.Arguments(0)) Then
    WScript.Echo "�t�H���_�ł͂���܂���"
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
'[�T�v]
' �t�H���_���C���v���V�[�W��
'*********************************************************
sub ReNameFolder(folder)
    Dim f
    Dim element
    For Each f In folder.subFolders
        For Each element In myDic
            Dim objRe: Set objRe = New RegExp
            ' �p�^�[��
            objRe.Pattern = element
            ' �S�̂�����
            objRe.Global = True
            ' �啶������������ʂ��Ȃ�
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
