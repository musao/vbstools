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
sub_import "clsCmArray.vbs"
sub_import "clsCmBroker.vbs"
sub_import "clsCmBufferedReader.vbs"
sub_import "clsCmBufferedWriter.vbs"
sub_import "clsCmCalendar.vbs"
sub_import "clsCmCharacterType.vbs"
sub_import "clsCmCssGenerator.vbs"
sub_import "clsCmHtmlGenerator.vbs"
sub_import "libCom.vbs"

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
    Set PoWriter = new_WriterTo(func_CM_FsGetPrivateLogFilePath, 8, True, -1)
    '�u���[�J�[�N���X�̃C���X�^���X�̐ݒ�
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe "log", GetRef("sub_GetPathsLogger")
    '�p�����[�^�i�[�p�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
    fw_excuteSub "sub_GetPathsGetParameters", oParams, oBroker
    
    '�����̃t�@�C���p�X���N���b�v�{�[�h�ɏo�͂���
    fw_excuteSub "sub_GetPathsProc", oParams, oBroker
    
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
    Dim oArg : Set oArg = fw_storeArguments()
    '�����O�o��
    sub_GetPathsLogger Array(9, "sub_GetPathsGetParameters", cf_toString(oArg))
    
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
    fs_writeFile sTempFilePaths, oParam.join(vbNewLine)
    new_Shell().Run "cmd /c clip <""" & sTempFilePaths & """", 0, True
    
    '�ꎞ�t�@�C�����폜
    fs_deleteFile sTempFilePaths
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GetPathsLogger()
'Overview                    : ���O�o�͂���
'Detailed Description        : fw_logger()�ɈϏ�����
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
    fw_logger avParams, PoWriter
End Sub
