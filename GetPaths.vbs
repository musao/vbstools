'***************************************************************************************************
'FILENAME                    : GetPaths.vbs
'Overview                    : �����̃t�@�C���p�X���N���b�v�{�[�h�ɃR�s�[����
'Detailed Description        : Sendto����g�p����
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
Private Const Cs_FOLDER_LIB = "lib"
Private PoWriter, PoBroker

'import��`
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
Call sub_import("clsCmArray.vbs")
Call sub_import("clsCmBufferedWriter.vbs")
Call sub_import("clsCmCalendar.vbs")
Call sub_import("clsCmBroker.vbs")
Call sub_import("clsCompareExcel.vbs")
Call sub_import("libCom.vbs")

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
    '���O�o�͂̐ݒ�
    Set PoWriter = new_WriterTo(func_CM_FsGetPrivateLogFilePath, 8, True, -2)
    '�o��-�w�ǌ^�iPublish/subscribe�j�C���X�^���X�̐ݒ�
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe "log", GetRef("sub_GetPathsLogger")
    '�p�����[�^�i�[�p�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
    sub_CM_ExcuteSub "sub_GetPathsGetParameters", oParams, oBroker
    
    '�����̃t�@�C���p�X���N���b�v�{�[�h�ɏo�͂���
    sub_CM_ExcuteSub "sub_GetPathsProc", oParams, oBroker
    
    '���O�o�͂��N���[�Y
    PoWriter.close()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set oBroker = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GetPathsGetParameters()
'Overview                    : ���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
'Detailed Description        : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g��Key="Param"�Ŋi�[����
'                              �z��iclsCmArray�^�j�ɖ��O�Ȃ������i/Key:Value �`���łȂ��j��S��
'                              �擾����
'Argument
'     aoParams               : �p�����[�^�i�[�p�I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetPathsGetParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '�����O�o��
    sub_GetPathsLogger Array(9, "sub_GetPathsGetParameters", func_CM_ToStringArguments())
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    cf_bindAt aoParams, "Param", oArg.Item("Unnamed").slice(0,vbNullString)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GetPathsProc()
'Overview                    : �����̃t�@�C���p�X���N���b�v�{�[�h�ɏo�͂���
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�I�u�W�F�N�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetPathsProc( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '�ꎞ�t�@�C���ɘA�������������o��
    Dim sTempFilePaths : sTempFilePaths = func_CM_FsGetTempFilePath() 
    sub_CM_FsWriteFile sTempFilePaths, oParam.join(vbNewLine)
'    sub_CM_FsWriteFile sTempFilePaths, oParam.joinVbs(vbNewLine)
    CreateObject("Wscript.Shell").Run "cmd /c clip <""" & sTempFilePaths & """", 0, True
    
    '�ꎞ�t�@�C�����폜
    func_CM_FsDeleteFile sTempFilePaths
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GetPathsLogger()
'Overview                    : ���O�o�͂���
'Detailed Description        : sub_CM_UtilLogger()�ɈϏ�����
'Argument
'     avParams               : �z��^�̃p�����[�^���X�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GetPathsLogger( _
    byRef avParams _
    )
    sub_CM_UtilLogger avParams, PoWriter
End Sub
