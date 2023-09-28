'***************************************************************************************************
'FILENAME                    : CompExcel.vbs
'Overview                    : �G�N�Z���t�@�C�����r����
'Detailed Description        : �H����
'Argument
'     PATH1                  : ��r����G�N�Z���t�@�C���̃p�X1
'     PATH2                  : ��r����G�N�Z���t�@�C���̃p�X2
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
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
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    '���O�o�͂̐ݒ�
    Dim sPath : sPath = func_CM_FsGetPrivateLogFilePath()
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    '�o��-�w�ǌ^�iPublish/subscribe�j�C���X�^���X�̐ݒ�
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_CmpExcelLogger"))
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v�錾
    Dim oParams : Set oParams = new_Dictionary()
    
    '������
    Call sub_CM_ExcuteSub("sub_CmpExcelInitialize", oParams, PoPubSub, "log")
    
    '���X�N���v�g�̈����擾
    Call sub_CM_ExcuteSub("sub_CmpExcelGetParameters", oParams, PoPubSub, "log")
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    Call sub_CM_ExcuteSub("sub_CmpExcelDispInputFiles", oParams, PoPubSub, "log")
    
    '�G�N�Z���t�@�C�����r����
    Call sub_CM_ExcuteSub("sub_CmpExcelCompareFiles", oParams, PoPubSub, "log")
    
    '�I������
    Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", oParams, PoPubSub, "log")
    
    '�t�@�C���ڑ����N���[�Y����
    PoWriter.FileClose
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoPubSub = Nothing
    Set PoWriter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_CmpExcelInitialize()
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
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelInitialize( _
    byRef aoParams _
    )
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : ���X�N���v�g�̈����擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="Param"�Ŋi�[����
'                              �p�����[�^�i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2)                 ��r����G�N�Z���t�@�C���̃p�X
'                              ���������葶�݂���t�@�C���p�X�̂ݎ擾����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelGetParameters( _
    byRef aoParams _
    )
    '�I���W�i���̈������擾
    Dim oArg : Set oArg = func_CM_UtilStoringArguments()
    '�����O�o��
    Call sub_CmpExcelLogger(Array(9, "sub_CmpExcelGetParameters", "Arguments are " & func_CM_ToStringArguments()))
    
    '�p�����[�^�i�[�p�I�u�W�F�N�g�ɐݒ�
    Call sub_CM_BindAt(aoParams, "Param", oArg.Item("Unnamed").Slice(0,2))
    
    Set oArg = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
'Function/Sub Name           : sub_CmpExcelDispInputFiles()
'Overview                    : ��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="Parameter"�Ŋi�[����
'                              �p�����[�^�i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2)                 ��r����G�N�Z���t�@�C���̃p�X
'                              ��r����G�N�Z���t�@�C���̃p�X��2�����̏ꍇ�ɕs�������擾�i�[����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelDispInputFiles( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    Const Cs_TITLE_EXCEL = "��r�Ώۃt�@�C�����J��"
    
    If oParam.Length > 1 Then
    '�p�����[�^��2�ȏゾ������֐��𔲂���
        Exit Sub
    End If
    
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParam.Length > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            '�t�@�C���I���L�����Z���̏ꍇ�͓��X�N���v�g���I������
                Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", aoParams, PoPubSub, "log")
                PoWriter.FileClose
                Wscript.Quit
            End If
            If func_CM_FsFileExists(sPath) Then
            '�t�@�C�������݂���ꍇ�p�����[�^���擾
                oParam.Push sPath
            End If
        Loop
        
        .Quit
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 4
'Function/Sub Name           : sub_CmpExcelCompareFiles()
'Overview                    : �G�N�Z���t�@�C�����r����
'Detailed Description        : �G���[�͖�������
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelCompareFiles( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    '4-1 ��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
    Call sub_CM_ExcuteSub("sub_CmpExcelSortByDateLastModified", aoParams, PoPubSub, "log")
    
    '4-2 ��r
    With New clsCompareExcel
        .PathFrom = oParam(0)
        .PathTo = oParam(1)
        .Compare()
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 4-1
'Function/Sub Name           : sub_CmpExcelSortByDateLastModified()
'Overview                    : ��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CmpExcelSortByDateLastModified( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    Dim oParam : Set oParam = aoParams.Item("Param")
    
    Dim oDateTimeA : Set oDateTimeA = new_clsCalSetDate(func_CM_FsGetFile(oParam(0)).DateLastModified)
    Dim oDateTimeB : Set oDateTimeB = new_clsCalSetDate(func_CM_FsGetFile(oParam(1)).DateLastModified)
    If oDateTimeA.CompareTo(oDateTimeB) > 0 Then
    '�ŏ��̃t�@�C���̕����V�����i�ŏI�X�V�����傫���j�ꍇ�A���Ԃ����ւ���
        oParam.Reverse
    End If
    
    '�I�u�W�F�N�g���J��
    Set oParam = Nothing
    Set oDateTimeA = Nothing
    Set oDateTimeB = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 5
'Function/Sub Name           : sub_CmpExcelTerminate()
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
Private Sub sub_CmpExcelTerminate( _
    byRef aoParams _
    )
    PoWriter.Flush
End Sub

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : sub_CmpExcelLogger()
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
Private Sub sub_CmpExcelLogger( _
    byRef avParams _
    )
    Call sub_CM_UtilCommonLogger(avParams, PoWriter)
End Sub
