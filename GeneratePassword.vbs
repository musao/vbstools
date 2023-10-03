'***************************************************************************************************
'FILENAME                    : GeneratePassword.vbs
'Overview                    : �p�X���[�h�𐶐�����
'Detailed Description        : ���������p�X���[�h�̓N���b�v�{�[�h�ɃR�s�[����
'Argument                    : �ȉ��̖��O�t�������i/Key:Value �`���j�̂݁A���O�Ȃ������͖�������
'                                /Length : ��������p�X���[�h�̕�����
'                                /U      : ��������p�X���[�h�̕�����ɔ��p�p���啶�����g�p����
'                                /L      : ��������p�X���[�h�̕�����ɔ��p�p�����������g�p����
'                                /N      : ��������p�X���[�h�̕�����ɔ��p�������g�p����
'                                /S      : ��������p�X���[�h�̕�����ɋL�����g�p����
'                                            �L���̎��   !"#$%&'()*+,-./:;<=>?[\]^_`{|}~�i31��ށj
'                                /Add    : �ǉ��w�肷�镶����i�J���}��؂�ŕ����w��\�j
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
Private PoWriter, PoPubSub

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
    '�p�����[�^�i�[�p�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dictionary()
    
    '���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
    Call sub_CM_ExcuteSub("sub_GnrtPwGetParameters", oParams, PoPubSub, "log")
    
    '�p�X���[�h�𐶐�����
    Call sub_CM_ExcuteSub("sub_GnrtPwGenerate", oParams, PoPubSub, "log")
    
    '�t�@�C���ڑ����N���[�Y����
    Call PoWriter.Close()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_GnrtPwGetParameters()
'Overview                    : ���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
'Detailed Description        : ���O�t�������i/Key:Value �`���j�������擾����
'                              Key           Value                                     Default
'                              ------------  ----------------------------------------  -------------
'                              "Param"       �p�����[�^�̉�͌���
'
'                              ���O�t�������i/Key:Value �`���j�̍\��
'                              Key           Value                                     Default
'                              ------------  ----------------------------------------  -------------
'                              "Length"      �����̒���                                16
'                                            �����̎��                                �S�Ċ܂�
'                               "U"           ���p�p���啶��
'                               "L"           ���p�p��������
'                               "N"           ���p����
'                               "S"           �S�Ă̋L��
'                              "Add"         �ǉ��w�肷�镶������J���}��؂�Ŏw��    �Ȃ�
'Argument
'     aoParams               : �p�����[�^�i�[�p�I�u�W�F�N�g
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
    '�����O�o��
    Call sub_GnrtPwLogger(Array(9, "sub_GnrtPwGetParameters", func_CM_ToStringArguments()))
    
    '�����̓��e�����
    
    '�����̒���
    Dim oKey, lLength
    oKey = "Length"
    If oArg.Item("Named").Exists(oKey) Then lLength = oArg.Item("Named").Item(oKey) Else lLength = 16
    
    '�ǉ��w�肷�镶����
    Dim vAdd
    oKey = "Add"
    If oArg.Item("Named").Exists(oKey) Then 
        vAdd = new_ArraySplit(oArg.Item("Named").Item(oKey), ",", vbBinaryCompare).Items
    Else
        vAdd = Empty
    End If
    
    '�����̎��
    Dim oSetting, lSum, lType
    Set oSetting = new_DictSetValues(Array("U", 1, "L", 2, "N", 4, "S", 8))
    lSum = 0
    For Each oKey In oSetting.Keys
        If oArg.Item("Named").Exists(oKey) Then lSum = lSum + oSetting.Item(oKey)
    Next
    lType = lSum
    If lType = 0 And func_CM_ArrayIsAvailable(vAdd)<>True Then lType = 15
    
    Dim oParam : Set oParam = new_DictSetValues(Array("Length", lLength, "Type", lType, "Additional", vAdd))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    Call sub_CM_BindAt(aoParams, "Param", oParam)
    
    Set oParam = Nothing
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_GnrtPwGenerate()
'Overview                    : �p�X���[�h�𐶐�����
'Detailed Description        : ���������p�X���[�h�̓N���b�v�{�[�h�ɃR�s�[���AInputBox�ɕ\������
'Argument
'     aoParams               : �p�����[�^�i�[�p�I�u�W�F�N�g
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
    '�p�X���[�h����
    Dim lLength, lType, vAdd
    With aoParams.Item("Param")
        Call sub_CM_Bind(lLength, .Item("Length"))
        Call sub_CM_Bind(lType, .Item("Type"))
        Call sub_CM_Bind(vAdd, .Item("Additional"))
    End With
    Dim sPw : sPw = func_CM_UtilGenerateRandomString(lLength, lType, vAdd)
    
    '�����O�o��
    Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "GeneratedPassword is " & sPw))
    
    '�_�C�A���O�̃��b�Z�[�W�Ȃǂ��쐬
    Dim sMsg, sTitle
    sMsg = "�p�X���[�h�𐶐����܂���" & vbNewLine & "OK�{�^������������ƃN���b�v�{�[�h�ɃR�s�[���܂�"
    sTitle = new_clsCalGetNow() & " �ɍ쐬"
    
    '�����O�o��
    Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "Display Inputbox."))
    '�ꎞ�t�@�C���̃p�X���쐬
    Dim sPath : sPath = func_CM_FsGetTempFilePath()
    Do Until Inputbox(sMsg, sTitle, sPw)=False
        '�ꎞ�t�@�C���ɐ��������p�X���[�h���o��
        Call sub_CM_FsWriteFile(sPath, sPw)
        '�N���b�v�{�[�h�Ɉꎞ�t�@�C���̓��e���o��
        Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sPath & """", 0, True)
        '�ꎞ�t�@�C�����폜
        Call func_CM_FsDeleteFile(sPath)
        '�����O�o��
        Call sub_GnrtPwLogger(Array(3, "sub_GnrtPwGenerate", "Copied to clipboard."))
    Loop
    
    
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
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
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilLogger(avParams, PoWriter)
End Sub
