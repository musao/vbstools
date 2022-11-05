'***************************************************************************************************
'FILENAME                    : clsFsBaseText.vbs
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
    Call func_clsFsBaseText_1(oUtAssistant)
    
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
'Function/Sub Name           : func_clsFsBaseText_1()
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
Private Sub func_clsFsBaseText_1( _
    byRef aoUtAssistant _
    )
    
    Call aoUtAssistant.Run("func_clsFsBaseText_1_1")
    Call aoUtAssistant.Run("func_clsFsBaseText_1_2")
    Call aoUtAssistant.Run("func_clsFsBaseText_1_3")
    Call aoUtAssistant.Run("func_clsFsBaseText_1_4")
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_clsFsBaseText_1_1()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i����L���b�V���Ȃ��j
'Detailed Description        : �H����
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
Private Function func_clsFsBaseText_1_1( _
    )
    Dim boFlg : boFlg = True
    
    '�ꎞ�t�@�C�����쐬
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    
    '���Ғl���擾
    Dim oExpect : Set oExpect = func_clsFsBaseTextGetExpectedValue(func_CM_FsGetFile(sPath))
    
    '�e�X�g�Ώێ��s
    Dim oSut : Set oSut = New clsFsBase
    
    With oSut
        '�L���b�V���g�p�Ŏ��s
        .UseCache = True
        .ValidPeriod = 3600
        .Path = sPath
        
        '����
        boFlg = func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        If .UseCache <> True Then boFlg = False
        If .ValidPeriod <> 3600 Then boFlg = False
        If .MostRecentReference = 0 Then boFlg = False
    End With
    
    '�ꎞ�t�@�C���̍폜
    Call func_CM_FsDeleteFile(sPath)
    
    func_clsFsBaseText_1_1 = boFlg
    
    Set oExpect = Nothing
    Set oSut = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : func_clsFsBaseText_1_2()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ڃL���b�V���g�p�������j
'Detailed Description        : �H����
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
Private Function func_clsFsBaseText_1_2( _
    )
    Dim boFlg : boFlg = True
    
    '�ꎞ�t�@�C�����쐬
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    
    '���Ғl���擾
    Dim oExpect : Set oExpect = func_clsFsBaseTextGetExpectedValue(func_CM_FsGetFile(sPath))
    
    '�e�X�g�Ώێ��s
    Dim oSut : Set oSut = New clsFsBase
    
    With oSut
        '�L���b�V���g�p�Ŏ��s
        .UseCache = True
        .ValidPeriod = 3600
        .Path = sPath
        
        '����
        Call func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        boFlg = func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        If .UseCache <> True Then boFlg = False
        If .ValidPeriod <> 3600 Then boFlg = False
        If .MostRecentReference = 0 Then boFlg = False
    End With
    
    '�ꎞ�t�@�C���̍폜
    Call func_CM_FsDeleteFile(sPath)
    
    func_clsFsBaseText_1_2 = boFlg
    
    Set oExpect = Nothing
    Set oSut = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : 1-3
'Function/Sub Name           : func_clsFsBaseText_1_3()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ڃL���b�V���g�p�����؂�j
'Detailed Description        : �H����
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
Private Function func_clsFsBaseText_1_3( _
    )
    Dim boFlg : boFlg = True
    
    '�ꎞ�t�@�C�����쐬
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    
    '���Ғl���擾
    Dim oExpect : Set oExpect = func_clsFsBaseTextGetExpectedValue(func_CM_FsGetFile(sPath))
    
    '�e�X�g�Ώێ��s
    Dim oSut : Set oSut = New clsFsBase
    
    With oSut
        '�L���b�V���g�p�Ŏ��s
        .UseCache = True
        .ValidPeriod = 0
        .Path = sPath
        
        '����
        Call func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        boFlg = func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        If .UseCache <> True Then boFlg = False
        If .ValidPeriod <> 0 Then boFlg = False
        If .MostRecentReference = 0 Then boFlg = False
    End With
    
    '�ꎞ�t�@�C���̍폜
    Call func_CM_FsDeleteFile(sPath)
    
    func_clsFsBaseText_1_3 = boFlg
    
    Set oExpect = Nothing
    Set oSut = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : 1-4
'Function/Sub Name           : func_clsFsBaseText_1_4()
'Overview                    : �e�v���p�e�B�̒l�̎擾�̐������i2��ڃL���b�V�������j
'Detailed Description        : �H����
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
Private Function func_clsFsBaseText_1_4( _
    )
    Dim boFlg : boFlg = True
    
    '�ꎞ�t�@�C�����쐬
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(sPath)
    If Not(func_CM_FsFileExists(sPath)) Then Exit Function
    
    '���Ғl���擾
    Dim oExpect : Set oExpect = func_clsFsBaseTextGetExpectedValue(func_CM_FsGetFile(sPath))
    
    '�e�X�g�Ώێ��s
    Dim oSut : Set oSut = New clsFsBase
    
    With oSut
        '�L���b�V���g�p�Ŏ��s
        .UseCache = False
        .ValidPeriod = 3600
        .Path = sPath
        
        '����
        Call func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        boFlg = func_clsFsBaseTextValidateAllItems(oSut, oExpect)
        If .UseCache <> False Then boFlg = False
        If .ValidPeriod <> 3600 Then boFlg = False
        If .MostRecentReference = 0 Then boFlg = False
    End With
    
    '�ꎞ�t�@�C���̍폜
    Call func_CM_FsDeleteFile(sPath)
    
    func_clsFsBaseText_1_4 = boFlg
    
    Set oExpect = Nothing
    Set oSut = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTextGetExpectedValue()
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
Private Function func_clsFsBaseTextGetExpectedValue( _
    byRef aoSomeObject _
    )
    
    Dim oExpect : Set oExpect = CreateObject("Scripting.Dictionary")
    With aoSomeObject
        Call oExpect.Add("Attributes", .Attributes)
        Call oExpect.Add("DateCreated", .DateCreated)
        Call oExpect.Add("DateLastAccessed", .DateLastAccessed)
        Call oExpect.Add("DateLastModified", .DateLastModified)
        Call oExpect.Add("Drive", .Drive)
        Call oExpect.Add("Name", .Name)
        Call oExpect.Add("ParentFolder", .ParentFolder)
        Call oExpect.Add("Path", .Path)
        Call oExpect.Add("ShortName", .ShortName)
        Call oExpect.Add("ShortPath", .ShortPath)
        Call oExpect.Add("Size", .Size)
        Call oExpect.Add("Type", .Type)
    End With
    
    Set func_clsFsBaseTextGetExpectedValue = oExpect
    Set oExpect = Nothing
    
End Function

'***************************************************************************************************
'Processing Order            : none
'Function/Sub Name           : func_clsFsBaseTextValidateAllItems()
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
Private Function func_clsFsBaseTextValidateAllItems( _
    byRef aoSut _
    , byRef aoExpect _
    )
    Dim boFlg : boFlg = True
    
    With aoSut
        If .Attributes <> aoExpect.Item("Attributes") Then boFlg = False
        If .DateCreated <> aoExpect.Item("DateCreated") Then boFlg = False
        If .DateLastAccessed <> aoExpect.Item("DateLastAccessed") Then boFlg = False
        If .DateLastModified <> aoExpect.Item("DateLastModified") Then boFlg = False
        If .Drive <> aoExpect.Item("Drive") Then boFlg = False
        If .Name <> aoExpect.Item("Name") Then boFlg = False
        If Not (.ParentFolder Is aoExpect.Item("ParentFolder")) Then boFlg = False
        If .Path <> aoExpect.Item("Path") Then boFlg = False
        If .ShortName <> aoExpect.Item("ShortName") Then boFlg = False
        If .ShortPath <> aoExpect.Item("ShortPath") Then boFlg = False
        If .Size <> aoExpect.Item("Size") Then boFlg = False
        If .FileFolderType <> aoExpect.Item("Type") Then boFlg = False
    End With
    
    func_clsFsBaseTextValidateAllItems = boFlg
    
End Function

