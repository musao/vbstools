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
    Set PoPubSub = Nothing
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
    Call PoWriter.FileClose()
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    Set PoWriter = Nothing
    Set PoPubSub = Nothing
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
    '���O�o�͂̐ݒ�
    Dim sPath : sPath = func_CM_FsBuildPath( _
                    func_CM_FsGetPrivateFolder("log") _
                    , func_CM_FsGetGetBaseName(WScript.ScriptName) & new_clsCalGetNow().DisplayFormatAs("_YYMMDD_hhmmss.000.log") _
                    )
    Set PoWriter = new_clsCmBufferedWriter(func_CM_FsOpenTextFile(sPath, 8, True, -2))
    
    '�o��-�w�ǌ^�iPublish/subscribe�j�C���X�^���X�̐ݒ�
    Set PoPubSub = new_clsCmPubSub()
    Call PoPubSub.Subscribe("log", GetRef("sub_CmpExcelLogger"))
'    Call sub_CM_BindAt( aoParams, "PubSub", oPubSub)
    
End Sub

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_CmpExcelGetParameters()
'Overview                    : ���X�N���v�g�̈����擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="Parameter"�Ŋi�[����
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
    '�p�����[�^�i�[�p�n�b�V���}�b�v
    Dim oParameter : Set oParameter = new_Dictionary()
    
    Dim lCnt : lCnt = 0
    Dim sParam
    For Each sParam In WScript.Arguments
        If func_CM_FsFileExists(sParam) Then
        '�t�@�C�������݂���ꍇ�p�����[�^���擾
            lCnt = lCnt + 1
            Call sub_CM_BindAt(oParameter, lCnt, sParam)
        End If
    Next
    
    Call sub_CM_BindAt(aoParams, "Parameter", oParameter)
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
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
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    Const Cs_TITLE_EXCEL = "��r�Ώۃt�@�C�����J��"
    
    If oParameter.Count > 1 Then
    '�p�����[�^��2�ȏゾ������֐��𔲂���
        Exit Sub
    End If
    
    With CreateObject("Excel.Application")
        Dim sPath
        Do Until oParameter.Count > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            '�t�@�C���I���L�����Z���̏ꍇ�͓��X�N���v�g���I������
                Call sub_CM_ExcuteSub("sub_CmpExcelTerminate", aoParams, "log")
                Wscript.Quit
            End If
            If func_CM_FsFileExists(sPath) Then
            '�t�@�C�������݂���ꍇ�p�����[�^���擾
                Call oParameter.Add(oParameter.Count+1, sPath)
            End If
        Loop
        
        .Quit
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
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
    
    '4-1 ��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
    Call sub_CM_ExcuteSub("sub_CmpExcelSortByDateLastModified", aoParams, PoPubSub, "log")
    
    '4-2 ��r
    With New clsCompareExcel
'        Call sub_CM_Bind(.PubSub, aoParams.Item("PubSub"))
        .PathFrom = aoParams.Item("Parameter").Item(1)
        .PathTo = aoParams.Item("Parameter").Item(2)
        .Compare()
    End With

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
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    With oParameter
        Dim oDateTimeA : Set oDateTimeA = new_clsCalSetDate(func_CM_FsGetFile(.Item(1)).DateLastModified)
        Dim oDateTimeB : Set oDateTimeB = new_clsCalSetDate(func_CM_FsGetFile(.Item(2)).DateLastModified)
        If oDateTimeA.CompareTo(oDateTimeB) <= 0 Then
        '�ŏ��̃t�@�C���̕����Â��i�ŏI�X�V�����������j�ꍇ�A�����𔲂���
            Exit Sub
        End If
        
        '�l�����ւ���
        Dim sValue1, sValue2
        sValue1 = .Item(1)
        sValue2 = .Item(2)
        
        .RemoveAll
        
        Call .Add(1, sValue2)
        Call .Add(2, sValue1)
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
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
'    '�t�@�C���ڑ����N���[�Y����
'    Call PoWriter.FileClose()
'    Set PoWriter = Nothing
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
    With PoWriter
        Dim sCont : sCont = new_clsCalGetNow()
        sCont = sCont & vbTab & avParams(0)
        sCont = sCont & vbTab & avParams(1)
        sCont = sCont & vbTab & avParams(2)
        .WriteContents(sCont)
        .newLine()
    End With
End Sub
