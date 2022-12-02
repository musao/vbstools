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
    Call sub_clsFsBaseTest_1(oUtAssistant)
    
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
Private Sub sub_clsFsBaseTest_1( _
    byRef aoUtAssistant _
    )
    
    Call sub_clsFsBaseTest_1_1(aoUtAssistant)
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseTest_1_1()
'Overview                    : clsFsBase�̑S�����̊m���炵�����m�F����
'Detailed Description        : �H����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1( _
    byRef aoUtAssistant _
    )
    
    Call sub_clsFsBaseTest_1_1_1(aoUtAssistant)
    Call sub_clsFsBaseTest_1_1_2(aoUtAssistant)
    Call sub_clsFsBaseTest_1_1_3(aoUtAssistant)
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-1
'Function/Sub Name           : sub_clsFsBaseTest_1_1_1()
'Overview                    : �e�����̎擾�̐��������m�F����
'Detailed Description        : ���{����
'                              �E�t�@�C��/�t�H���_���ꂼ��ɂ��Č��؂���
'                              �E�L���b�V���g�p�ۂ͔�
'                              �E�L���b�V���L�����Ԃ�0�b
'                              �E�S�����̒l��1��擾
'                              ���Ғl
'                              �E�S�����̒l������������
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
Private Sub sub_clsFsBaseTest_1_1_1( _
    byRef aoUtAssistant _
    )
    '�t�@�C��/�t�H���_�̃p�^�[��
    Dim boTargetIsFile : Dim boTargetIsFiles(1)
    boTargetIsFiles(0) = True : boTargetIsFiles(1) = False
    '�e�����̃p�^�[��
    Dim sPropNames
    sPropNames = Array("Attributes", "DateCreated", "DateLastAccessed", "DateLastModified", "Drive", _
        "Name", "ParentFolder", "Path", "ShortName", "ShortPath", "Size", "Type")
    
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim lCntOut : Dim lCntIn : Dim sPropName
    
    '�t�@�C��/�t�H���_���ꂼ��ɂ���
    Dim oVariationTargetIsFile()
    For lCntOut=0 To Ubound(boTargetIsFiles)
        boTargetIsFile = boTargetIsFiles(lCntOut)
        
        Call sub_CM_ArrayAddItem(oVariationTargetIsFile, func_clsFsBaseTest_1_1_1_C(sPropName, boTargetIsFile, sPropNames))
        
    Next
    Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_1_", oVariationTargetIsFile)
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-1
'Function/Sub Name           : sub_clsFsBaseTest_1_1_1()
'Overview                    : �e�����̎擾�̐��������m�F����
'Detailed Description        : ���{����
'                              �E�t�@�C��/�t�H���_���ꂼ��ɂ��Č��؂���
'                              �E�L���b�V���g�p�ۂ͔�
'                              �E�L���b�V���L�����Ԃ�0�b
'                              �E�S�����̒l��1��擾
'                              ���Ғl
'                              �E�S�����̒l������������
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
Private Function func_clsFsBaseTest_1_1_1_C( _
    byVal asSubName _
    , byVal aboTargetIsFile _
    , byVal asPropName _
    )
    Set func_clsFsBaseTest_1_1_1_C = _
        func_clsFsBaseTestCreateArgumentFor_1_1_x(asSubName, aboTargetIsFile, False, 0, asPropName, False, False)
End Function
'***************************************************************************************************
'Processing Order            : 1-1-2
'Function/Sub Name           : sub_clsFsBaseTest_1_1_2()
'Overview                    : �L���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ��m�F����
'Detailed Description        : ���{����
'                              �E�t�@�C��/�t�H���_���ꂼ��ɂ��Č��؂���
'                              �E�L���b�V���L�����Ԃ�0�b
'                              �E�C�ӂ̑����̒l��1��擾
'                              ���Ғl
'                              �E�L���b�V���g�p�ۂ��ς���Ă��Ȃ�����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_2( _
    byRef aoUtAssistant _
    )
    '�t�@�C��/�t�H���_�̃p�^�[��
    Dim boTargetIsFile : Dim boTargetIsFiles(2)
    boTargetIsFiles(1) = True : boTargetIsFiles(1) = False
    '�L���b�V���g�p�ۂ̃p�^�[��
    Dim boUseCaches
    boUseCaches = Array(True, False)
    
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim sPropName : sPropName = "Attributes"
    Dim lCntOut : Dim lCntIn : Dim boUseCache
    
    '�t�@�C��/�t�H���_���ꂼ��ɂ���
    For lCntOut=1 To Ubound(boTargetIsFiles)
        boTargetIsFile = boTargetIsFiles(lCntOut)
        oPatterns.RemoveAll
        
        '�L���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ̌���
        For lCntIn = 0 To Ubound(boUseCaches)
            boUseCache = boUseCaches(lCntIn)
            oPatterns.Add boUseCache, func_clsFsBaseTestCreateArgumentFor_1_1_x(boTargetIsFile, boUseCache, 0, sPropName, True, False)
        Next
        Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_2_" & Cstr(lCntOut) & "_", oPatterns)
    Next
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-3
'Function/Sub Name           : sub_clsFsBaseTest_1_1_3()
'Overview                    : �L���b�V���L�����ԁi�b���j���ς���Ă��Ȃ����Ƃ��m�F����
'Detailed Description        : ���{����
'                              �E�t�@�C��/�t�H���_���ꂼ��ɂ��Č��؂���
'                              �E�L���b�V���g�p�ۂ͉�
'                              �E�C�ӂ̑����̒l��1��擾
'                              ���Ғl
'                              �E�L���b�V���L�����ԁi�b���j���ς���Ă��Ȃ�����
'Argument
'     aoUtAssistant          : �P�̃e�X�g�p�A�V�X�^���g�N���X�̃C���X�^���X
'Return Value
'     ���� True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_clsFsBaseTest_1_1_3( _
    byRef aoUtAssistant _
    )
    '�t�@�C��/�t�H���_�̃p�^�[��
    Dim boTargetIsFile : Dim boTargetIsFiles(2)
    boTargetIsFiles(1) = True : boTargetIsFiles(1) = False
    '�L���b�V���L�����ԁi�b���j�̃p�^�[��
    Dim lValidPeriods
    lValidPeriods = Array(0,1,2147483647,-1,-2147483648)
    
    Dim oPatterns : Set oPatterns = CreateObject("Scripting.Dictionary")
    Dim sPropName : sPropName = "Attributes"
    Dim lCntOut : Dim lCntIn : Dim lValidPeriod
    
    '�t�@�C��/�t�H���_���ꂼ��ɂ���
    For lCntOut=1 To Ubound(boTargetIsFiles)
        boTargetIsFile = boTargetIsFiles(lCntOut)
        oPatterns.RemoveAll
        
        '�t�@�C���̑����̌���
        For lCntIn = 0 To Ubound(lValidPeriods)
            lValidPeriod = lValidPeriods(lCntIn)
            oPatterns.Add Cstr(lValidPeriod), func_clsFsBaseTestCreateArgumentFor_1_1_x(True, True, lValidPeriod, sPropName, False, True)
        Next
        Call aoUtAssistant.RunWithMultiplePatterns("func_clsFsBaseTest_1_1_3_" & Cstr(lCntOut) & "_", oPatterns)
    Next
    
    Set oPatterns = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1-x
'Function/Sub Name           : func_clsFsBaseTestCreateArgumentFor_1_1_x()
'Overview                    : func_clsFsBaseTest_1_1()�p�̈������n�b�V���}�b�v���쐬
'Detailed Description        : func_clsFsBaseTestCreateArgument()���Q��
'Argument
'     asSubName              : �P�[�X�̃T�u����
'     aboTargetIsFile        : �Ώۂ̓t�@�C�����ۂ�
'     aboUseCache            : �L���b�V���g�p��
'     alValidPeriod          : �L���b�V���L�����ԁi�b���j
'     asPropName1            : 1��ڂɎ擾���鑮�����i2��ڂ��Ȃ��ꍇ�͒l�����؂���j
'     aboDontChgUc           : �Ō�ɃL���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
'     aboDontChgVp           : �Ō�ɃL���b�V���L�����ԁi�b���j���ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
'Return Value
'     �������n�b�V���}�b�v
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgumentFor_1_1_x( _
    byVal asSubName _
    , byVal aboTargetIsFile _
    , byVal aboUseCache _
    , byVal alValidPeriod _
    , byVal asPropName1 _
    , byVal aboDontChgUc _
    , byVal aboDontChgVp _
    )
    Set func_clsFsBaseTestCreateArgumentFor_1_1_x = _
        func_clsFsBaseTestCreateArgument( _
            asSubName _
            , aboTargetIsFile _
            , aboUseCache _
            , alValidPeriod _
            , False _
            , False _
            , False _
            , 0 _
            , asPropName1 _
            , vbNullString _
            , aboDontChgUc _
            , aboDontChgVp _
            , vbNullString _
            , vbNullString _
            )
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestCreateArgument()
'Overview                    : �P�[�X�p�^�[�����i�[�p�n�b�V���}�b�v�ɓo�^����������n�b�V���}�b�v���쐬����
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "Conditions"             ���{�����̃n�b�V���}�b�v
'                              "Inspections"            ���ؓ��e�̃n�b�V���}�b�v
'
'                              ���{�����̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "TargetIsFile"           �Ώۂ̓t�@�C�����ۂ�
'                              "UseCache"               �L���b�V���g�p��
'                              "ValidPeriod"            �L���b�V���L�����ԁi�b���j
'                              "SetFsoFlg"              FileSystemObject�I�u�W�F�N�g�̐ݒ�v��
'                              "DoItTwice"              �����擾��2�񂷂邩�ۂ�
'                              "IsRecreate"             2��ڂ̑����擾�̒��O�ɑΏۃt�@�C��/�t�H���_���č쐬���邩�ۂ�
'                              "SleepMSecond"           1��ڂ̑����擾�̒���ɃX���[�v���鎞�ԁi�~���b�j
'
'                              ���ؓ��e�̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName1"              1��ڂɎ擾���鑮�����i2��ڂ��Ȃ��ꍇ�͒l�����؂���j
'                              "PropName2"              2��ڂɎ擾���鑮�����A�l�����؂���
'                              "DontChgUc"              �Ō�ɃL���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
'                              "DontChgVp"              �Ō�ɃL���b�V���L�����ԁi�b���j���ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
'                              "IsUpdLcct"              �ŏI�L���b�V���m�F���Ԃ��Ō�̑����擾�̒��O����ς���Ă��邩�ۂ�
'                              "IsUpdLcut"              �ŏI�L���b�V���X�V���Ԃ��Ō�̑����擾�̒��O����ς���Ă��邩�ۂ�
'Argument
'     asSubName              : �P�[�X�̃T�u����
'     aboTargetIsFile        : ���{�����̃n�b�V���}�b�v��"TargetIsFile"�Ɠ���
'     aboUseCache            : ���{�����̃n�b�V���}�b�v��"UseCache"�Ɠ���
'     alValidPeriod          : ���{�����̃n�b�V���}�b�v��"ValidPeriod"�Ɠ���
'     aboSetFsoFlg           : ���{�����̃n�b�V���}�b�v��"SetFsoFlg"�Ɠ���
'     aboDoItTwice           : ���{�����̃n�b�V���}�b�v��"DoItTwice"�Ɠ���
'     aboIsRecreate          : ���{�����̃n�b�V���}�b�v��"IsRecreate"�Ɠ���
'     alSleepMSecond         : ���{�����̃n�b�V���}�b�v��"SleepMSecond"�Ɠ���
'     asPropName1            : ���ؓ��e�̃n�b�V���}�b�v��"PropName1"�Ɠ���
'     asPropName2            : ���ؓ��e�̃n�b�V���}�b�v��"PropName2"�Ɠ���
'     aboDontChgUc           : ���ؓ��e�̃n�b�V���}�b�v��"DontChgUc"�Ɠ���
'     aboDontChgVp           : ���ؓ��e�̃n�b�V���}�b�v��"DontChgVp"�Ɠ���
'     aboIsUpdLcct           : ���ؓ��e�̃n�b�V���}�b�v��"IsUpdLcct"�Ɠ���
'     aboIsUpdLcut           : ���ؓ��e�̃n�b�V���}�b�v��"IsUpdLcut"�Ɠ���
'Return Value
'     �������n�b�V���}�b�v
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_clsFsBaseTestCreateArgument( _
    byVal asSubName _
    , byVal aboTargetIsFile _
    , byVal aboUseCache _
    , byVal alValidPeriod _
    , byVal aboSetFsoFlg _
    , byVal aboDoItTwice _
    , byVal aboIsRecreate _
    , byVal alSleepMSecond _
    , byVal asPropName1 _
    , byVal asPropName2 _
    , byVal aboDontChgUc _
    , byVal aboDontChgVp _
    , byVal aboIsUpdLcct _
    , byVal aboIsUpdLcut _
    )
    Dim oConditions : Set oConditions = CreateObject("Scripting.Dictionary")
    With oConditions
        .Add "TargetIsFile", aboTargetIsFile
        .Add "UseCache", aboUseCache
        .Add "ValidPeriod", alValidPeriod
        .Add "SetFsoFlg", aboSetFsoFlg
        .Add "DoItTwice", aboDoItTwice
        .Add "IsRecreate", aboIsRecreate
        .Add "SleepMSecond", alSleepMSecond
    End With
    
    Dim oInspections : Set oInspections = CreateObject("Scripting.Dictionary")
    With oInspections
        .Add "PropName1", asPropName1
        .Add "PropName2", asPropName2
        .Add "DontChgUc", aboDontChgUc
        .Add "DontChgVp", aboDontChgVp
        .Add "IsUpdLcct", aboIsUpdLcct
        .Add "IsUpdLcut", aboIsUpdLcut
    End With
    
    Dim oArgument : Set oArgument = CreateObject("Scripting.Dictionary")
    With oArgument
        .Add "SubName", asSubName
        .Add "Conditions", oConditions
        .Add "Inspections", oInspections
    End With
    
    Set func_clsFsBaseTestCreateArgument = oArgument
    Set oInspections = Nothing
    Set oConditions = Nothing
    Set oArgument = Nothing
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : ����
'Overview                    : �P�[�X���ƂɃm�[�}���P�[�X�ėp���s�ɈϏ�����֐�
'Detailed Description        : func_clsFsBaseTestCreateArgumentFor_x_x()���Q��
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
Private Function func_clsFsBaseTest_1_1_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_1_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_1_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_1_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_2_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_2_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_2_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_2_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_3_1_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_3_1_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function
Private Function func_clsFsBaseTest_1_1_3_2_( _
    byRef aoArgument _
    )
    func_clsFsBaseTest_1_1_3_2_ = func_clsFsBaseTestNormalBase(aoArgument)
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTestNormalBase()
'Overview                    : �m�[�}���P�[�X�ėp���s
'Detailed Description        : �������n�b�V���}�b�v�̍\����func_clsFsBaseTestCreateArgument()���Q��
'                              �{�֐��Ŏg�p���鍀�ڂɌ��肵�ċL�ڂ���
'                              ���{�����̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "TargetIsFile"           �Ώۂ̓t�@�C�����ۂ�
'                              "UseCache"               �L���b�V���g�p��
'                              "ValidPeriod"            �L���b�V���L�����ԁi�b���j
'
'                              ���ؓ��e�̃n�b�V���}�b�v�̓��e
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "PropName1"              1��ڂɎ擾���鑮�����i2��ڂ��Ȃ��ꍇ�͒l�����؂���j
'                              "DontChgUc"              �Ō�ɃL���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
'                              "DontChgVp"              �Ō�ɃL���b�V���L�����ԁi�b���j���ς���Ă��Ȃ����Ƃ����؂��邩�ۂ�
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
Private Function func_clsFsBaseTestNormalBase( _
    byRef aoArgument _
    )
    '�������̎擾
    With aoArgument.Item("Conditions")
    '���{����
        Dim boTargetIsFile : boTargetIsFile = .Item("TargetIsFile")
        Dim boUseCache : boUseCache = .Item("UseCache")
        Dim dbValidPeriod : dbValidPeriod = .Item("ValidPeriod")
    End With
    With aoArgument.Item("Inspections")
    '���ؓ��e
        Dim sPropName : sPropName = .Item("PropName1")
        Dim boDontChgUc : boDontChgUc = .Item("DontChgUc")
        Dim boDontChgVp : boDontChgVp = .Item("DontChgVp")
    End With
    
    '�O���� �ꎞ�t�@�C��/�t�H���_�쐬�A���Ғl�擾
    Dim oExpect
    Dim boResult : boResult = True
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    If boTargetIsFile Then
        Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
        If Not(func_CM_FsFileExists(sPath)) Then Exit Function
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFile(sPath))
    Else
        Call func_CM_FsCreateFolder(sPath)
        If Not(func_CM_FsFolderExists(sPath)) Then Exit Function
        Set oExpect = func_clsFsBaseTestGetExpectedValue(func_CM_FsGetFolder(sPath))
    End If
    
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
        
        '�L���b�V���g�p�ۂ��ς���Ă��Ȃ����Ƃ̌���
        If (boDontChgUc=True) Then boResult = (boUseCache = .UseCache)
        
        '�L���b�V���L�����ԁi�b���j���ς���Ă��Ȃ����Ƃ̌���
        If (boDontChgVp=True) Then boResult = (dbValidPeriod = .ValidPeriod)
        
    End With
    
    '�㏈�� �ꎞ�t�@�C��/�t�H���_�폜
    If boTargetIsFile Then Call func_CM_FsDeleteFile(sPath) Else Call func_CM_FsDeleteFolder(sPath)
    Set oExpect = Nothing
    
    '���ʕԋp
    func_clsFsBaseTestNormalBase = boResult
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

