'*********************************************************
' ���C������
'*********************************************************
Option Explicit
Const STR_SCH_PATH = "C:\work\qiita\VBS\FlieCopy\�R�s�[��"
Const STR_DEST_PATH = "C:\work\qiita\VBS\FlieCopy\�R�s�[��"
Const STR_TARGET_FIlENAME = "�Ώ�"
Dim objFileSys: Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
CALL FolderCopy(objFileSys, STR_SCH_PATH, STR_DEST_PATH)
set objFileSys = Nothing

msgbox "end"

'*********************************************************
'[�T�v]
' �R�s�[���{����t�@���N�V����
'*********************************************************
Function IsTarget(objFile)
    If objFile.Name = STR_TARGET_FIlENAME Then
        IsTarget = True
    Else
        IsTarget = False
    End If
End Function
'*********************************************************
'[�T�v]
' �t�@�C���R�s�[�v���V�[�W���[
'*********************************************************
Sub FolderCopy(objFileSys, strSchPath, strDestPath)
    Dim objFolder 
    Dim objFile
    Dim objSubFolder
    Set objFolder = objFileSys.GetFolder(strSchPath)
    IF IsTarget(objFile) Then   
        For Each objFile in objFolder.Files
            If Not objFileSys.FolderExists(StrDestPath) Then
                CALL CreateDirectory(objFileSys, StrDestPath)
            End If
            objFileSys.CopyFile ObjFile.Path, StrDestPath & "/" & objFile.Name
        Next
    End If
    For Each ObjSubFolder in objFolder.subFolders
        CALL FolderCopy(objFileSys, objSubFolder.Path, strDestPath & "\" & ObjSubFolder.Name)
    Next
End Sub
'*********************************************************
'[�T�v]
' �t�@�C���R�s�[�v���V�[�W���[
'*********************************************************
Sub CreateDirectory(objFileSys, strPath)
    Dim strParentFolder
    strParentFolder = objFileSys.GetParentFolderName(strPath)
    If Not objFileSys.FolderExists(strParentFolder) Then
        CALL CreateDirectory(objFileSys, StrPath)
    End If
    If Not objFileSys.FolderExists(StrPath) Then
        objFileSys.CreateFolder StrPath
    End If
End Sub