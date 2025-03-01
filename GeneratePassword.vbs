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
'                                            �L���̎��   !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~�i32��ށj
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

'�ϐ�
Private PoWriter

'lib import
Private Const Cs_FOLDER_LIB = "lib"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_LIB)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing

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
    Set PoWriter = new_WriterTo(fw_getLogPath, 8, True, -1)
    '�u���[�J�[�N���X�̃C���X�^���X�̐ݒ�
    Dim oBroker : Set oBroker = new_Broker()
    oBroker.subscribe topic.LOG, GetRef("sub_GnrtPwLogger")
    '�p�����[�^�i�[�p�I�u�W�F�N�g�錾
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈������p�����[�^�i�[�p�I�u�W�F�N�g�Ɏ擾����
    fw_excuteSub "sub_GnrtPwGetParameters", oParams, oBroker
    
    '�p�X���[�h�𐶐�����
    fw_excuteSub "sub_GnrtPwGenerate", oParams, oBroker
    
    '���O�o�͂��N���[�Y
    PoWriter.close()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set oBroker = Nothing
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
    Dim oArg : Set oArg = fw_storeArguments()
    '�����O�o��
    sub_GnrtPwLogger Array(logType.DETAIL, "sub_GnrtPwGetParameters", cf_toString(oArg))
    
    '�����̓��e�����
    
    '�����̒���
    Dim oKey, lLength
    oKey = "Length"
    If oArg.Item("Named").Exists(oKey) Then lLength = oArg.Item("Named").Item(oKey) Else lLength = 16
    
    '�ǉ��w�肷�镶����
    Dim vAdd
    oKey = "Add"
    If oArg.Item("Named").Exists(oKey) Then 
        vAdd = new_ArrSplit(oArg.Item("Named").Item(oKey), ",", vbBinaryCompare).toArray()
    Else
        vAdd = Empty
    End If
    
    '�����̎��
    Dim oSetting, lSum, lType
    Set oSetting = new_DicOf(Array("U", 1, "L", 2, "N", 4, "S", 8))
    lSum = 0
    For Each oKey In oSetting.Keys
        If oArg.Item("Named").Exists(oKey) Then lSum = lSum + oSetting.Item(oKey)
    Next
    lType = lSum
    If lType = 0 And IsEmpty(vAdd) Then lType = 15
    
    Dim oParam : Set oParam = new_DicOf(Array("Length", lLength, "Type", lType, "Additional", vAdd))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    cf_bindAt aoParams, "Param", oParam
    
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
        cf_bind lLength, .Item("Length")
        cf_bind lType, .Item("Type")
        cf_bind vAdd, .Item("Additional")
    End With
    Dim vCharList : vCharList = new_Char().charList(lType)
    vCharList = Filter(vCharList, " ", False, vbBinaryCompare)
    If Not IsEmpty(vAdd) Then cf_pushA vCharList, vAdd
    Dim sPw : sPw = util_randStr(vCharList, lLength)
    
    '�����O�o��
    sub_GnrtPwLogger Array(logType.INFO, "sub_GnrtPwGenerate", "GeneratedPassword is " & sPw)
    
    '�_�C�A���O�̃��b�Z�[�W�Ȃǂ��쐬
    Dim sMsg, sTitle
    sMsg = "�p�X���[�h�𐶐����܂���" & vbNewLine & "OK�{�^������������ƃN���b�v�{�[�h�ɃR�s�[���܂�"
    sTitle = new_Now() & " �ɍ쐬"
    
    '�����O�o��
    sub_GnrtPwLogger Array(logType.INFO, "sub_GnrtPwGenerate", "Display Inputbox.")
    '�ꎞ�t�@�C���̃p�X���쐬
    Dim sPath : sPath = fw_getTempPath()
    Do Until Inputbox(sMsg, sTitle, sPw)=False
        '�ꎞ�t�@�C���ɐ��������p�X���[�h���o��
        fs_writeFileDefault sPath, sPw
        '�N���b�v�{�[�h�Ɉꎞ�t�@�C���̓��e���o��
        fw_runShellSilently "cmd /c clip <" & fs_wrapInQuotes(sPath)
        '�ꎞ�t�@�C�����폜
        fs_deleteFile sPath
        '�����O�o��
        sub_GnrtPwLogger Array(logType.INFO, "sub_GnrtPwGenerate", "Copied to clipboard.")
    Loop
    
    
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_GnrtPwLogger()
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
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_GnrtPwLogger( _
    byRef avParams _
    )
    fw_logger avParams, PoWriter
End Sub
