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

'lib\com import
Dim sRelativeFolderName : sRelativeFolderName = "lib\com"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing
'lib import
sRelativeFolderName = "lib"
With new_FSO()
    sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    ExecuteGlobal .OpenTextfile(.BuildPath(sLibFolderPath,"libEnum.vbs")).ReadAll
End With


'���O�o�͐�A�u���[�J�[�N���X�̃C���X�^���X�̐ݒ�
Private PoTs4Log, PoBroker
Set PoTs4Log = fw_getTextstreamForLog()
Set PoBroker = new_BrokerOf(Array(topic.LOG, GetRef("this_logger")))

'Main�֐����s
Call Main()

'�I������
PoTs4Log.close()
Set PoBroker = Nothing : Set PoTs4Log = Nothing
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
    '�p�����[�^�i�[�p�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
    fw_excuteSub "this_getParameters", oParams, PoBroker
    
    '�����̃t�@�C���p�X���N���b�v�{�[�h�ɏo�͂���
    fw_excuteSub "this_toClipbord", oParams, PoBroker
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
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
Private Sub this_getParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = fw_storeArguments()
    '�����O�o��
    this_logger Array(logType.DETAIL, "this_getParameters()", cf_toString(oArg))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    cf_bindAt aoParams, "Param", new_ArrOf(oArg.Item("Unnamed")).slice(0,vbNullString)
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : this_toClipbord()
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
Private Sub this_toClipbord( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '�ꎞ�t�@�C���ɘA�������������o��
    Dim sTempFilePaths : sTempFilePaths = fw_getTempPath()
    fs_writeFileDefault sTempFilePaths, this_replaceEnvironmentStrings(oParam.join(vbNewLine))
    fw_runShellSilently "cmd /c clip <" & fs_wrapInQuotes(sTempFilePaths)
    
    '�ꎞ�t�@�C�����폜
    fs_deleteFile sTempFilePaths
    
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : this_replaceEnvironmentStrings()
'Overview                    : ���ϐ��ɒu��������
'Detailed Description        : �H����
'Argument
'     asStr                  : �Ώ�
'Return Value
'     ���ϐ��ɒu��������������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/04/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function this_replaceEnvironmentStrings( _
    byVal asStr _
    )
    Dim sSettings
    sSettings = Array("%UserProfile%")

    Dim sRet : sRet = asStr
    Dim i
    For Each i In sSettings
        sRet = Replace(sRet, new_Shell().ExpandEnvironmentStrings(i), i)
    Next

    this_replaceEnvironmentStrings = sRet
End Function

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : this_logger()
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
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoTs4Log
End Sub
