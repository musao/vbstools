'***************************************************************************************************
'FILENAME                    : GetPath.vbs
'Overview                    : �����̃t�@�C���p�X���N���b�v�{�[�h�ɃR�s�[����
'Detailed Description        : Sendto����t�@�C���p�X���擾����̂Ɏg�p����
'Argument
'     PATH1,2...             : �t�@�C���̃p�X1,2,...
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_FOLDER_TEMP = "tmp"

'Include�p�֐���`
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'Include
Call sub_Include("VbsBasicLibCommon.vbs")


'���C���֐����s
Call Main()
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : ���C���֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim sLineFeedCode : sLineFeedCode = vbCrLf
    
    '���������s�ŘA������
    Dim sOutput : sOutput = ""
    Dim sItem
    For Each sItem In Wscript.Arguments
        If Not(Len(sItem)) Then
            If (Len(sOutput)) Then sOutput = sOutput & sLineFeedCode
            sOutput = sOutput & sItem
        End If
    Next
    
    '�ꎞ�t�@�C���̃p�X���쐬
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
    Dim sFolderPath : sFolderPath = func_CM_FsBuildPath(sParentFolderPath, Cs_FOLDER_TEMP)
    If Not(func_CM_FsFolderExists(sFolderPath)) Then func_CM_FsCreateFolder(sFolderPath)
    Dim sTempFilePaths : sTempFilePaths = func_CM_FsBuildPath(sFolderPath, func_CM_FsGetTempFileName())
    
    '�ꎞ�t�@�C���ɘA�������������o��
    Call sub_CM_FsWriteFile(sTempFilePaths, sOutput)
    
    '�N���b�v�{�[�h�Ɉꎞ�t�@�C���̓��e���o��
    Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sTempFilePaths & """", 0, True)
    
    '�ꎞ�t�@�C�����폜
    Call func_CM_FsDeleteFile(sTempFilePaths)
    
End Sub
