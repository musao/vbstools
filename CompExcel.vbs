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
Call sub_Include("VbsBasicLibCommon.vbs")
Call sub_Include("clsCompareExcel.vbs")


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
    
    Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
    
    '���X�N���v�g�̈����擾
    Call sub_CmpExcelGetParameters( _
                            oParams _
                             )
    
    '��r�Ώۃt�@�C�����͉�ʂ̕\���Ǝ擾
    Call sub_CmpExcelDispInputFiles( _
                            oParams _
                             )
    
    '�G�N�Z���t�@�C�����r����
    Call sub_CmpExcelCompareFiles( _
                            oParams _
                             )
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : 1
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
    Dim oParameter : Set oParameter = CreateObject("Scripting.Dictionary")
    Dim lCnt : lCnt = 0
    Dim sParam
    For Each sParam In WScript.Arguments
        If func_CM_FsFileExists(sParam) Then
        '�t�@�C�������݂���ꍇ�p�����[�^���擾
            lCnt = lCnt + 1
            Call oParameter.Add(lCnt, sParam)
        End If
    Next
    
    Call aoParams.Add("Parameter", oParameter)
End Sub

'***************************************************************************************************
'Processing Order            : 2
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
    
    Dim oExcel : Set oExcel = CreateObject("Excel.Application")
    With oExcel
        Dim sPath
        Do Until oParameter.Count > 1
            
            sPath = .GetOpenFilename( , , Cs_TITLE_EXCEL, , False)
            If sPath = False Then
            '�t�@�C���I���L�����Z���̏ꍇ�͓��X�N���v�g���I������
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
    Set oExcel = Nothing
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 3
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
    
    '3-1 ��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
    Call sub_CmpExcelSortByDateLastModified(aoParams)
    
    '3-2 ��r
    With New clsCompareExcel
        .PathFrom = aoParams.Item("Parameter").Item(1)
        .PathTo = aoParams.Item("Parameter").Item(2)
        .Compare()
    End With

End Sub

'***************************************************************************************************
'Processing Order            : 3-1
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
    
    If func_CM_FsGetFile(oParameter.Item(1)).DateLastModified _
        <= _
        func_CM_FsGetFile(oParameter.Item(2)).DateLastModified _
        Then
    '�ŏ��̃t�@�C���̕����Â��i�ŏI�X�V�����������j�ꍇ�A�����𔲂���
        Exit Sub
    End If
    
    '�l�����ւ���
    With oParameter
        Dim sValue1 : Dim sValue2
        sValue1 = .Item(1)
        sValue2 = .Item(2)
        
        .RemoveAll
        
        Call .Add(1, sValue2)
        Call .Add(2, sValue1)
    End With
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
End Sub
