'***************************************************************************************************
'FILENAME                    : clsFsBaseTest.vbs
'Overview                    : �t�@�C���E�t�H���_���ʃN���X�̃e�X�g
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_UTLIB_FILE = "VbsUtLib.vbs"
Private Const Cs_UTAST_FILE = "clsUtAssistant.vbs"
Private Const Cs_COMMON_FILE = "VbsBasicLibCommon.vbs"
Private Const Cs_TEST_FILE = "clsFsBase.vbs"

With CreateObject("Scripting.FileSystemObject")
    '�P�̃e�X�g�p���C�u�����ǂݍ���
    Dim sIncludeFolderPath : sIncludeFolderPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTLIB_FILE)).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTAST_FILE)).ReadAll
    '���ʃ��C�u�����ǂݍ���
    sIncludeFolderPath = .BuildPath(.GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName)), Cs_FOLDER_INCLUDE)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_COMMON_FILE)).ReadAll
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
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim oUtAssistant : Set oUtAssistant = New clsUtAssistant
    
    '�m�[�}���P�[�X�̃e�X�g
    Call func_clsFsBaseTest_1(oUtAssistant)
    
    '���ʏo��
    Call sub_UtResultOutput(oUtAssistant)
    
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
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_OutputReport( _
    byRef aoUtAssistant _
    )
    Call sub_UtWriteFile(func_UtGetThisLogFilePath(), aoUtAssistant.OutputReportInTsvFormat())
End Sub


'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : func_clsFsBaseTest_1()
'Overview                    : �m�[�}���P�[�X�̃e�X�g
'Detailed Description        : �H����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_clsFsBaseTest_1( _
    byRef aoUtAssistant _
    )
    
    Call func_clsFsBaseTest_1_1(aoUtAssistant)
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_1")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_2")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_3")
'    Call aoUtAssistant.Run("func_clsFsBaseTest_1_4")
'    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i1��ځj
'Detailed Description        : ���{����
'                              �E�L���b�V���g�p�ۂ͉�
'                              �E�L���b�V���L�����Ԃ�3600�b
'                              �E�S�v���p�e�B�̒l��1��擾
'                              ���Ғl
'                              �E�S�v���p�e�B�̒l������������
'                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
'                              �E�L���b�V���m�F����i�ŏI�L���b�V���m�F���Ԃ������l�łȂ����Ɓj
'                              �E�L���b�V���g�p�Ȃ��i�ŏI�L���b�V���X�V���Ԃ������l�łȂ����Ɓj
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_clsFsBaseTest_1_1( _
    byRef aoUtAssistant _
    )
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim lNum : lNum = 0
    Dim sPropName
    
    lNum = lNum + 1 : sPropName = "Attributes" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateCreated" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateLastAccessed" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "DateLastModified" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Drive" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Name" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ParentFolder" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Path" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ShortName" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "ShortPath" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Size" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    lNum = lNum + 1 : sPropName = "Type" : oPatterns.Add lNum & "_" & sPropName, func_clsFsBaseTest_1_1_CreateArgument(sPropName)
    
    Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_", oPatterns)
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i1��ځj
'Detailed Description        : ���{����
'                              �E�L���b�V���g�p�ۂ͉�
'                              �E�L���b�V���L�����Ԃ�3600�b
'                              �E�S�v���p�e�B�̒l��1��擾
'                              ���Ғl
'                              �E�S�v���p�e�B�̒l������������
'                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
'                              �E�L���b�V���m�F����i�ŏI�L���b�V���m�F���Ԃ������l�łȂ����Ɓj
'                              �E�L���b�V���g�p�Ȃ��i�ŏI�L���b�V���X�V���Ԃ������l�łȂ����Ɓj
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_1_CreateArgument( _
    byVal asPropName _
    )
    Dim oArgument : Set oArgument = CreateObject("Scripting.Dictionary")
    Dim oConditions : Set oConditions = CreateObject("Scripting.Dictionary")
    Dim oInspections : Set oInspections = CreateObject("Scripting.Dictionary")
    
    oConditions.Add "UseCache", False
    oConditions.Add "ValidPeriod", 0
    
    oInspections.Add "PropName", asPropName
    
    oArgument.Add "Conditions", oConditions
    oArgument.Add "Inspections", oInspections
    
    Set func_clsFsBaseTest_1_1_CreateArgument = oArgument
    
    Set oInspections = Nothing
    Set oConditions = Nothing
    Set oArgument = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : func_clsFsBaseTest_1_1_()
'Overview                    : �����Ŏw�肵�������̒l�̐��������m�F����
'Detailed Description        : �������̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "Conditions"             ���{�����̃n�b�V���}�b�v
'                              "Inspections"            ���ؓ��e�̃n�b�V���}�b�v
'
'                              ���{�����̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "UseCache"               �L���b�V���g�p��
'                              "ValidPeriod"            �L���b�V���L�����ԁi�b���j
'
'                              ���ؓ��e�̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName"               ���ؑΏۂ̑�����
'
'Argument
'     aoArgument             : �������̃n�b�V���}�b�v
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_1_( _
    byRef aoArgument _
    )
    '�������̎擾
    With aoArgument.Item("Conditions")
    '���{����
        Dim boUseCache : boUseCache = .Item("UseCache")
        Dim dbValidPeriod : dbValidPeriod = .Item("ValidPeriod")
    End With
    With aoArgument.Item("Inspections")
    '���ؓ��e
        Dim sPropName : sPropName = .Item("PropName")
    End With
    
    Dim boResult : boResult = True
    
    '�ꎞ�t�@�C���쐬�A���Ғl�擾
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
    
    With New clsFsBase
        '�e�X�g�ΏۃN���X�ɏ������w��
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '�w�肵���v���p�e�B�̒l������
        If IsObject(oExpect.Item(sPropName)) Then
            If Not (.Prop(sPropName) Is oExpect.Item(sPropName)) Then boResult = False
        Else
            If .Prop(sPropName) <> oExpect.Item(sPropName) Then boResult = False
        End If
    End With
    
    '�ꎞ�t�@�C���폜
    Call func_CM_FsDeleteFile(sPath)
    
    '���{����
    func_clsFsBaseTest_1_1_ = boResult
    Set oExpect = Nothing
End Function

''***************************************************************************************************
''Processing Order            : 1-1
''Function/Sub Name           : func_clsFsBaseTest_1_1()
''Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i1��ځj
''Detailed Description        : ���{����
''                              �E�L���b�V���g�p�ۂ͉�
''                              �E�L���b�V���L�����Ԃ�3600�b
''                              �E�S�v���p�e�B�̒l��1��擾
''                              ���Ғl
''                              �E�S�v���p�e�B�̒l������������
''                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
''                              �E�L���b�V���m�F����i�ŏI�L���b�V���m�F���Ԃ������l�łȂ����Ɓj
''                              �E�L���b�V���g�p�Ȃ��i�ŏI�L���b�V���X�V���Ԃ������l�łȂ����Ɓj
''Argument
''     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
''Return Value
''     ���� True,Flase
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/11/03         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_clsFsBaseTest_1_1( _
'    )
'    Dim boResult : boResult = True
'    
'    '���{����
'    Dim boUseCache : boUseCache = True
'    Dim dbValidPeriod : dbValidPeriod = 3600
'    
'    '�e�X�g�Ώ�
'    Dim oSut : Set oSut = New clsFsBase
'    With oSut
'        '�ꎞ�t�@�C���쐬�A���Ғl�擾
'        Dim sPath : sPath = func_UtGetThisTempFilePath()
'        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
'        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
'        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
'        
'        '�e�X�g�ΏۃN���X�ɏ������w��
'        .UseCache = boUseCache
'        .ValidPeriod = dbValidPeriod
'        .Path = sPath
'        
'        '�S�v���p�e�B�̒l���擾�i1��ځj
'        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
'        
'        '����
'        If .UseCache <> boUseCache Then boResult = False
'        If .ValidPeriod <> dbValidPeriod Then boResult = False
'        If .LastCacheConfirmationTime = 0 Then boResult = False
'        If .LastCacheUpdateTime = 0 Then boResult = False
'        
'        '�ꎞ�t�@�C���폜
'        Call func_CM_FsDeleteFile(sPath)
'    End With
'    
'    '���{����
'    func_clsFsBaseTest_1_1 = boResult
'    Set oExpect = Nothing
'    Set oSut = Nothing
'End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : func_clsFsBaseTest_1_2()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ځA�L���b�V�������j
'Detailed Description        : ���{����
'                              �E�L���b�V���g�p�ۂ͔�
'                              �E�L���b�V���L�����Ԃ�3600�b
'                              �E�S�v���p�e�B�̒l��2��擾
'                              ���Ғl
'                              �E2��ڂɎ擾�����S�v���p�e�B�̒l������������
'                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
'                              �E�L���b�V���m�F�Ȃ��i�ŏI�L���b�V���m�F���Ԃ�1��ڎ擾�ォ��ς���Ă��Ȃ����Ɓj
'                              �E�L���b�V���g�p�Ȃ��i�ŏI�L���b�V���X�V���Ԃ�1��ڎ擾�ォ��ς���Ă��邱�Ɓj
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_2( _
    )
    Dim boResult : boResult = True
    
    '���{����
    Dim boUseCache : boUseCache = False
    Dim dbValidPeriod : dbValidPeriod = 3600
    
    '�e�X�g�Ώ�
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '�ꎞ�t�@�C���쐬�A���Ғl�擾
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        '�e�X�g�ΏۃN���X�ɏ������w��
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '�S�v���p�e�B�̒l���擾�i1��ځj
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10ms�X���[�v
        WScript.Sleep 10
        
        '�S�v���p�e�B�̒l���擾�i2��ځj
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '����
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime <> lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
        
        '�ꎞ�t�@�C���폜
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '���{����
    func_clsFsBaseTest_1_2 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-3
'Function/Sub Name           : func_clsFsBaseTest_1_3()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ځA�L���b�V���L�����Ԓ��߂��t�@�C���X�V�Ȃ��j
'Detailed Description        : ���{����
'                              �E�L���b�V���g�p�ۂ͉�
'                              �E�L���b�V���L�����Ԃ�0�b
'                              �E�S�v���p�e�B�̒l��2��擾
'                              �E1��ڂ�2��ڂŃt�@�C���̍ŏI�X�V�����ς���Ă��Ȃ�
'                              ���Ғl
'                              �E2��ڂɎ擾�����S�v���p�e�B�̒l������������
'                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
'                              �E�L���b�V���m�F����i�ŏI�L���b�V���m�F���Ԃ�1��ڎ擾�ォ��ς���Ă��邱�Ɓj
'                              �E�L���b�V���g�p����i�ŏI�L���b�V���X�V���Ԃ�1��ڎ擾�ォ��ς���Ă��Ȃ����Ɓj
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_3( _
    )
    Dim boResult : boResult = True
    
    '���{����
    Dim boUseCache : boUseCache = True
    Dim dbValidPeriod : dbValidPeriod = 0
    
    '�e�X�g�Ώ�
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '�ꎞ�t�@�C���쐬�A���Ғl�擾
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        '�e�X�g�ΏۃN���X�ɏ������w��
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '�S�v���p�e�B�̒l���擾�i1��ځj
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10ms�X���[�v
        WScript.Sleep 10
        
        '�S�v���p�e�B�̒l���擾�i2��ځj
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '����
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime <> lLastCacheUpdateTime Then boResult = False
        
        '�ꎞ�t�@�C���폜
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '���{����
    func_clsFsBaseTest_1_3 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 1-4
'Function/Sub Name           : func_clsFsBaseTest_1_4()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ځA�L���b�V���L�����Ԓ��߂��t�@�C���X�V����j
'Detailed Description        : ���{����
'                              �E�L���b�V���g�p�ۂ͉�
'                              �E�L���b�V���L�����Ԃ�0�b
'                              �E�S�v���p�e�B�̒l��2��擾
'                              �E1��ڂ�2��ڂŃt�@�C���̍ŏI�X�V�����ς���Ă��Ȃ�
'                              ���Ғl
'                              �E2��ڂɎ擾�����S�v���p�e�B�̒l������������
'                              �E�L���b�V���g�p�ہA���L�����Ԃ��ς��Ȃ�����
'                              �E�L���b�V���m�F����i�ŏI�L���b�V���m�F���Ԃ�1��ڎ擾�ォ��ς���Ă��邱�Ɓj
'                              �E�L���b�V���g�p�Ȃ��i�ŏI�L���b�V���X�V���Ԃ�1��ڎ擾�ォ��ς���Ă��邱�Ɓj
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTest_1_4( _
    )
    Dim boResult : boResult = True
    
    '���{����
    Dim boUseCache : boUseCache = True
    Dim dbValidPeriod : dbValidPeriod = 0
    
    '�e�X�g�Ώ�
    Dim oSut : Set oSut = New clsFsBase
    With oSut
        '�ꎞ�t�@�C���쐬�A���Ғl�擾
        Dim sPath : sPath = func_UtGetThisTempFilePath()
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Dim oExpect : Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        '�e�X�g�ΏۃN���X�ɏ������w��
        .UseCache = boUseCache
        .ValidPeriod = dbValidPeriod
        .Path = sPath
        
        '�S�v���p�e�B�̒l���擾�i1��ځj
        Call func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        Dim lLastCacheConfirmationTime : lLastCacheConfirmationTime = .LastCacheConfirmationTime
        Dim lLastCacheUpdateTime : lLastCacheUpdateTime = .LastCacheUpdateTime
        
        '10ms�X���[�v
        WScript.Sleep 10
        
        '�ꎞ�t�@�C���폜���č쐬�A���Ғl�̎擾
        Call func_CM_FsDeleteFile(sPath)
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        oExpect.RemoveAll
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
        
        '�S�v���p�e�B�̒l���擾�i2��ځj
        boResult = func_clsFsBaseTestValidateAllItems(oSut, oExpect)
        
        '����
        If .UseCache <> boUseCache Then boResult = False
        If .ValidPeriod <> dbValidPeriod Then boResult = False
        If .LastCacheConfirmationTime = lLastCacheConfirmationTime Then boResult = False
        If .LastCacheUpdateTime = lLastCacheUpdateTime Then boResult = False
        
        '�ꎞ�t�@�C���폜
        Call func_CM_FsDeleteFile(sPath)
    End With
    
    '���{����
    func_clsFsBaseTest_1_4 = boResult
    Set oExpect = Nothing
    Set oSut = Nothing
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestGetExpectedValue()
'Overview                    : ���Ғl�̎擾
'Detailed Description        : �H����
'Argument
'     aoSomeObject           : File/Folder�I�u�W�F�N�g
'Return Value
'     ���Ғl�̃n�b�V���}�b�v
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestGetExpectedValue( _
    byRef aoSomeObject _
    )
    
    Dim oExpect : Set oExpect = CreateObject("Scripting.Dictionary")
    With aoSomeObject
        oExpect.Add "Attributes", .Attributes
        oExpect.Add "DateCreated", .DateCreated
        oExpect.Add "DateLastAccessed", .DateLastAccessed
        oExpect.Add "DateLastModified", .DateLastModified
        oExpect.Add "Drive", .Drive
        oExpect.Add "Name", .Name
        oExpect.Add "ParentFolder", .ParentFolder
        oExpect.Add "Path", .Path
        oExpect.Add "ShortName", .ShortName
        oExpect.Add "ShortPath", .ShortPath
        oExpect.Add "Size", .Size
        oExpect.Add "Type", .Type
    End With
    
    Set func_clsFsBaseTestGetExpectedValue = oExpect
    Set oExpect = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestValidateAllItems()
'Overview                    : �S���ڂ̌��؂��s��
'Detailed Description        : �H����
'Argument
'     aoSut                  : �e�X�g�ΏۃN���X
'     aoExpect               : ���Ғl�̃n�b�V���}�b�v
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestValidateAllItems( _
    byRef aoSut _
    , byRef aoExpect _
    )
    Dim boFlg : boFlg = True
    
    With aoExpect
        Dim sKey
        For Each sKey In .Keys
            If IsObject(.Item(sKey)) Then
                If Not (aoSut.Prop(sKey) Is .Item(sKey)) Then boFlg = False
            Else
                If aoSut.Prop(sKey) <> .Item(sKey) Then boFlg = False
            End If
        Next
    End With
    
    func_clsFsBaseTestValidateAllItems = boFlg
    
End Function

