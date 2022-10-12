'***************************************************************************************************
'FILENAME                    : VbsBasicLibCommonTest.vbs
'Overview                    : ���ʊ֐����C�u�����̃e�X�g
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_UTLIB_FILE = "VbsUtLib.vbs"
Private Const Cs_UTAST_FILE = "clsUtAssistant.vbs"
Private Const Cs_TEST_FILE = "VbsBasicLibCommon.vbs"

With CreateObject("Scripting.FileSystemObject")
    '�P�̃e�X�g�p���C�u�����ǂݍ���
    Dim sIncludeFolderPath : sIncludeFolderPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTLIB_FILE)).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTAST_FILE)).ReadAll
    '�P�̃e�X�g�Ώۃ\�[�X�ǂݍ���
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName))
    sIncludeFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_TEST_FILE)).ReadAll
End With

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
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim oUtAssistant : Set oUtAssistant = New clsUtAssistant
    
    'func_CM_FsDeleteFile()�̃e�X�g
    Call func_CM_FsDeleteFileTest(oUtAssistant)
    'func_CM_FsGetParentFolderPath()�̃e�X�g
    Call func_CM_FsGetParentFolderPathTest(oUtAssistant)
    
    'UT���|�[�g�̏o��
    Call sub_OutputReport(oUtAssistant)
    
    '���ʂ����b�Z�[�W�ŏo��
    Dim sMsg : sMsg = "NG������܂��A���O���m�F��������"
    If oUtAssistant.isAllOk Then sMsg = "�S�P�[�XOK�ł��I"
    Call Msgbox(sMsg)
    
    Set oUtAssistant = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : Last
'Function/Sub Name           : sub_OutputReport()
'Overview                    : ���ʏo��
'Detailed Description        : �H����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_OutputReport( _
    byRef aoUtAssistant _
    )
    Call sub_UtWriteFile(func_UtGetThisLogFilePath(), aoUtAssistant.OutputReportInTsvFormat())
End Sub


'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : func_CM_FsDeleteFileTest()
'Overview                    : func_CM_FsDeleteFile()�̃e�X�g
'Detailed Description        : �H����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_CM_FsDeleteFileTest( _
    byRef aoUtAssistant _
    )
    
    '1-1 �폜����
    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestSuccess")
    '1-2 �폜���s
    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestFailure")
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_CM_FsDeleteFileTestSuccess()
'Overview                    : func_CM_FsDeleteFile()�̃e�X�g
'Detailed Description        : �폜�����̏ꍇ
'Argument
'     �Ȃ�
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFileTestSuccess( _
    )
    func_CM_FsDeleteFileTestSuccess = False
    
    '�ꎞ�t�@�C���̃t���p�X���擾
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    
    With CreateObject("Scripting.FileSystemObject")
        '�t�@�C�����쐬
        Call .CreateTextFile(sPath)
        
        '�t�@�C�����ł��Ă��邱�Ƃ��m�F
        If Not(.FileExists(sPath)) Then Exit Function
        
        '�e�X�g�Ώێ��s
        Dim boResult : boResult = func_CM_FsDeleteFile(sPath)
        
        '�߂�l���m�F
        If Not(boResult) Then Exit Function
        
        '�t�@�C�����폜�ł��Ă����琬��
        func_CM_FsDeleteFileTestSuccess = Not(.FileExists(sPath))
    End With
    
End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : func_CM_FsDeleteFileTestFailure()
'Overview                    : func_CM_FsDeleteFile()�̃e�X�g
'Detailed Description        : �폜���s�̏ꍇ
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFileTestFailure( _
    )
    func_CM_FsDeleteFileTestFailure = False
    
    '�ꎞ�t�@�C���̃t���p�X���擾
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    
    With CreateObject("Scripting.FileSystemObject")
        
        '�t�@�C�����Ȃ����Ƃ��m�F
        If .FileExists(sPath) Then Exit Function
        
        '�e�X�g�Ώێ��s
        Dim boResult : boResult = func_CM_FsDeleteFile(sPath)
        
        '�t�@�C�����Ȃ��̂Ŏ��s�����琬��
        func_CM_FsDeleteFileTestFailure = Not(boResult)
    End With
    
End Function

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : func_CM_FsGetParentFolderPathTest()
'Overview                    : func_CM_FsGetParentFolderPath()�̃e�X�g
'Detailed Description        : �H����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_CM_FsGetParentFolderPathTest( _
    byRef aoUtAssistant _
    )
    
    '2-1 �e�t�H���_������ꍇ
    Call aoUtAssistant.Run("func_CM_FsGetParentFolderPathTestNormal")
'    '1-2 �폜���s
'    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestFailure")
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : func_CM_FsGetParentFolderPathTestNormal()
'Overview                    : func_CM_FsGetParentFolderPath()�̃e�X�g
'Detailed Description        : �e�t�H���_������ꍇ
'Argument
'     �Ȃ�
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPathTestNormal( _
    )
    func_CM_FsGetParentFolderPathTestNormal = False
    
    '�ꎞ�t�@�C���̃t���p�X���擾
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    
    With CreateObject("Scripting.FileSystemObject")
        '�e�t�H���_�̃t���p�X���擾
        Dim sExpect : sExpect = .GetParentFolderName(sPath)
        
        '�e�X�g�Ώێ��s
        Dim sResult : sResult = func_CM_FsGetParentFolderPath(sPath)
        
        '���Ғl�Ɣ�r����
        func_CM_FsGetParentFolderPathTestNormal = (sExpect = sResult)
        
    End With
    
End Function