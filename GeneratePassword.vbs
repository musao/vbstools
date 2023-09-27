'***************************************************************************************************
'FILENAME                    : GeneratePassword.vbs
'Overview                    : �p�X���[�h�𐶐�����
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private PoWriter
Private PoPubSub

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
Call sub_Include("clsCmArray.vbs")
Call sub_Include("clsCmBufferedWriter.vbs")
Call sub_Include("clsCmCalendar.vbs")
Call sub_Include("clsCmPubSub.vbs")
Call sub_Include("clsCompareExcel.vbs")
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
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    '���O�o�͂̐ݒ�
    Dim sPath : sPath = func_CM_FsGetPrivateLogFilePath()
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    '�o��-�w�ǌ^�iPublish/subscribe�j�C���X�^���X�̐ݒ�
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_GnrtPwLogger"))
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v�錾
    Dim oParams : Set oParams = new_Dictionary()
    
    '������
    Call sub_CM_ExcuteSub("sub_GnrtPwInitialize", oParams, PoPubSub, "log")
    
    '���X�N���v�g�̈����擾
    Call sub_CM_ExcuteSub("sub_GnrtPwGetParameters", oParams, PoPubSub, "log")
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    Call sub_CM_ExcuteSub("sub_GnrtPwGenerate", oParams, PoPubSub, "log")
    
    '�I������
    Call sub_CM_ExcuteSub("sub_GnrtPwTerminate", oParams, PoPubSub, "log")
    
    '�t�@�C���ڑ����N���[�Y����
    Call PoWriter.FileClose()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GnrtPwInitialize()
'Overview                    : ������
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwInitialize( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GnrtPwGetParameters()
'Overview                    : ���X�N���v�g�̈����擾
'Detailed Description        : �����͂���������O�t�������i/Key:Value �`���j
'                              Key         Value                                        Default
'                              ----------  -------------------------------------------  -------------
'                              "Length"    �����̒���                                   16
'                              U,L,N,S     �����̎��                                   �S�Ċ܂�
'                               "U"         ���p�p���啶��
'                               "L"         ���p�p��������
'                               "N"         ���p����
'                               "S"         �S�Ă̋L��
'                              "Add"       �ǉ��w�肷�镶������J���}��؂�Ŏw��       �Ȃ�
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGetParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    Call sub_CM_BindAt(aoParams, "Arguments", oArg)
    
    '�����̒����ialLength�j
    Dim oKey, lLength
    oKey = "Length"
    If oArg.Item("Named").Exists(oKey) Then lLength = oArg.Item("Named").Item(oKey) Else lLength = 16
    
    '�ǉ��w�肷�镶����iavAdditional�j
    Dim vAdd
    oKey = "Add"
    If oArg.Item("Named").Exists(oKey) Then 
        vAdd = new_ArraySplit(oArg.Item("Named").Item(oKey), ",", vbBinaryCompare).Items
    Else
        vAdd = Empty
    End If
    
    '�����̎�ށialType�j
    Dim oSetting, lSum, lType
    Set oSetting = new_DictSetValues(Array("U", 1, "L", 2, "N", 4, "S", 8))
    lSum = 0
    For Each oKey In oSetting.Keys
        If oArg.Item("Named").Exists(oKey) Then lSum = lSum + oSetting.Item(oKey)
    Next
    lType = lSum
    If lType = 0 And func_CM_ArrayIsAvailable(vAdd)<>True Then lType = 15
    
    Dim oParam : Set oParam = new_DictSetValues(Array("Length", lLength, "Type", lType, "Additional", vAdd))
    Call sub_CM_BindAt(aoParams, "Parameter", oParam)
    
    Set oParam = Nothing
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_GnrtPwGenerate()
'Overview                    : �p�X���[�h����
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwGenerate( _
    byRef aoParams _
    )
    '�p�����[�^�̎擾
    Dim lLength, lType, vAdd
    With aoParams.Item("Parameter")
        Call sub_CM_Bind(lLength, .Item("Length"))
        Call sub_CM_Bind(lType, .Item("Type"))
        Call sub_CM_Bind(vAdd, .Item("Additional"))
    End With
    
    '�p�X���[�h����
    Dim sPw : sPw = func_CM_UtilGenerateRandomString(lLength, lType, vAdd)
    Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "GeneratedPassword is " & sPw))
    
    '�_�C�A���O�̃��b�Z�[�W�Ȃǂ��쐬
    Dim sMsg, sTitle
    sMsg = "�p�X���[�h�𐶐����܂���" & vbNewLine & "OK�{�^������������ƃN���b�v�{�[�h�ɃR�s�[���܂�"
    sTitle = new_clsCalGetNow() & " �ɍ쐬"
    
    '�ꎞ�t�@�C���̃p�X���쐬
    Dim sPath : sPath = func_CM_FsGetTempFilePath()
    Do
        '�ꎞ�t�@�C���ɐ��������p�X���[�h���o��
        Call sub_CM_FsWriteFile(sPath, sPw)
        '�N���b�v�{�[�h�Ɉꎞ�t�@�C���̓��e���o��
        Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sPath & """", 0, True)
        '�ꎞ�t�@�C�����폜
        Call func_CM_FsDeleteFile(sPath)
    Loop Until Inputbox(sMsg, sTitle, sPw)=False
    
    
End Sub

'***************************************************************************************************
'Processing Order            : 4
'Function/Sub Name           : sub_GnrtPwTerminate()
'Overview                    : �I������
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwTerminate( _
    byRef aoParams _
    )
    PoWriter.Flush
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
'Overview                    : ���O�o�͂���
'Detailed Description        : �H����
'Argument
'     avParams               : �z��^�̃p�����[�^���X�g
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilCommonLogger(avParams, PoWriter)
End Sub
