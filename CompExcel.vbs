'***************************************************************************************************
'FILENAME                    :CompExcel.vbs
'Generato                    :2017/04/26
'Descrition                  :�G�N�Z���t�@�C�����r����
' �p�����[�^�i�����j:
'     PATH         :�t�@�C���̃p�X
'---------------------------------------------------------------------------------------------------
'Modification Histroy
'
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         EXA Y.Fujii              Initial Release
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_FOLDER_TEMP = "tmp"

'Include�p�֐���`
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderName : sParentFolderName = oFso.GetParentFolderName(WScript.ScriptFullName)
    Dim sIncludeFilePath
    sIncludeFilePath = oFso.BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
    sIncludeFilePath = oFso.BuildPath(sIncludeFilePath, asIncludeFileName)
    ExecuteGlobal oFso.OpenTextfile(sIncludeFilePath).ReadAll
    Set oFso = Nothing
End Sub
'Include
Call sub_Include("VbsBasicLibCommon.vbs")

Main

Wscript.Quit

Sub Main()
    
    Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
    
    '�@�p�����[�^�̎擾
    Call sub_CmpExcelGetParameters( _
                            oParams _
                             )
    
    '�A��r�Ώۃt�@�C�����͉�ʂ̕\��
    Call sub_CmpExcelDispInputFiles( _
                            oParams _
                             )
    
    '�B�t�@�C�����r����
    Call sub_CmpExcelCompareFiles( _
                            oParams _
                             )
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    
End Sub

'�@�p�����[�^�̎擾
Private Sub sub_CmpExcelGetParameters( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�n�b�V���}�b�v
    Dim oParameter : Set oParameter = CreateObject("Scripting.Dictionary")
    Dim lCnt : lCnt = 0
    Dim sParam
    For Each sParam In WScript.Arguments
        If func_CM_FileExists(sParam) Then
        '�t�@�C�������݂���ꍇ�p�����[�^���擾
            lCnt = lCnt + 1
            Call oParameter.Add(lCnt, sParam)
        End If
    Next
    
    Call aoParams.Add("Parameter", oParameter)
End Sub

'�A��r�Ώۃt�@�C�����͉�ʂ̕\��
Private Sub sub_CmpExcelDispInputFiles( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�n�b�V���}�b�v
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
            If func_CM_FileExists(sPath) Then
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

'�B�t�@�C�����r����
Private Sub sub_CmpExcelCompareFiles( _
    byRef aoParams _
    )
 '   On Error Resume Next
    
    Dim oExcel : Set oExcel = CreateObject("Excel.Application")
    With oExcel
        .DisplayAlerts = False
        .ScreenUpdating = False
        .AutomationSecurity = 3                  'msoAutomationSecurityForceDisable = 3
    End With
    
    Dim lThemeColor(2)
    lThemeColor(1) = 2                           '�W�F 1(xlThemeColorLight1)
    lThemeColor(2) = 8                           '���� 4(xlThemeColorAccent4)
    
    '��r���ʗp�̐V�K���[�N�u�b�N���쐬
    Dim oWorkbookForResults : Set oWorkbookForResults = oExcel.Workbooks.Add(-4167)      '�V�K���[�N�u�b�N xlWBATWorksheet=-4167
    
    '�B�|�P�D��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
    Call sub_CmpExcelSortByDateLastModified(aoParams)
    
    '�B�|�Q�D��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
    Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(aoParams, oWorkbookForResults)
    
    '�B�|�R�D��r����
    Call sub_CmpExcelCompare(aoParams, oWorkbookForResults)

    '�I�u�W�F�N�g���J��
    Set oExcel = Nothing
    
End Sub

'�B�|�P�D��r����t�@�C�����Â����i�ŏI�X�V�������j�ɕ��בւ���
Private Sub sub_CmpExcelSortByDateLastModified( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�n�b�V���}�b�v
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    If func_CM_GetFile(oParameter.Item(1)).DateLastModified _
        <= _
        func_CM_GetFile(oParameter.Item(2)).DateLastModified _
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

'�B�|�Q�D��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
Private Sub sub_CmpExcelCopyAllSheetsToWorkbookForResults( _
    byRef aoParams _
    , byRef aoWorkbookForResults _
    )
    '�p�����[�^�i�[�p�n�b�V���}�b�v
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    '���[�N�V�[�g���Ƃ̃V�[�g�����l�[�����i�[�p�n�b�V���}�b�v
    Dim oWorkSheetRenameInfos : Set oWorkSheetRenameInfos = CreateObject("Scripting.Dictionary")
    
    Dim sPath : Dim sFromToString
    '��r���t�@�C���̃R�s�[
    sPath = oParameter.Item(1) : sFromToString = "From" 
    Call oWorkSheetRenameInfos.Add(sFromToString, _
        func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

    '��r��t�@�C���̃R�s�[
    sPath = oParameter.Item(2) : sFromToString = "To"
    Call oWorkSheetRenameInfos.Add(sFromToString, _
        func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))
    
    '���[�N�V�[�g���Ƃ̃V�[�g�����l�[�������i�[
    Call aoParams.Add("WorkSheetRenameInfos", oWorkSheetRenameInfos)

    aoWorkbookForResults.parent.ScreenUpdating = true
    aoWorkbookForResults.parent.visible = true
    stop

    '�I�u�W�F�N�g���J��
    Set oWorkSheetRenameInfos = Nothing
    Set oParameter = Nothing
End Sub

'�B�|�Q�|�P�D��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[�̏ڍ�
Private Function func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail( _
    byRef aoWorkbookForResults _
    , byVal asPath _
    , byVal asFromToString _
    )

    '��r�Ώۃt�@�C�����J��
    Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
    Dim oWorkBook : Set oWorkBook = func_CM_ExcelOpenFile(oExcel, asPath)
    Dim sTempPath : sTempPath = ""
    If oWorkBook.HasVBProject Then
    '�}�N������̏ꍇ�͕ʖ��ŕۑ�������ōēx�J���j
        sTempPath = func_CmpExcelGetTempFilePath()
        Call sub_CM_ExcelSaveAs(oWorkBook, sTempPath, vbNullString)
        Set oWorkBook = func_CM_ExcelOpenFile( oExcel, sTempPath)
    End If

    '�����̕ی����������
    Call sub_CM_OfficeUnprotect(oWorkBook, vbNullString)
    
    With oWorkBook
        '���[�N�V�[�g�̃��l�[�����i�[�p�n�b�V���}�b�v��`
        Dim oWorkSheetRenameInfo : Set oWorkSheetRenameInfo = CreateObject("Scripting.Dictionary")
        '�^�u�̐F�ϊ��p�n�b�V���}�b�v��`
        Dim oStringToThemeColor : Set oStringToThemeColor = CreateObject("Scripting.Dictionary")
        Call oStringToThemeColor.Add("From",2)
        Call oStringToThemeColor.Add("To",8)

        Dim oWorksheet
        Dim lCnt : lCnt = 0
        For Each oWorksheet In .Worksheets
            If oWorksheet.Visible Then
            '�S�Ă̌�����V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
                
                '�ύX�O��̃��[�N�V�[�g�����擾
                lCnt = lCnt + 1
                Call oWorkSheetRenameInfo.Add( _
                                lCnt, func_CmpExcelGetMapWorkSheetRenameInfo( _
                                                        oWorksheet.Name _
                                                        , func_CmpExcelMakeSheetName( _
                                                                                lCnt _
                                                                                , asFromToString _
                                                                                ) _
                                                        ) _
                                )
                
                '�V�[�g�̕\���𐮂���
                If oWorksheet.AutoFilterMode Then
                '�I�[�g�t�B���^���ݒ肳��Ă������������
                     oWorksheet.Cells(1,1).AutoFilter
                End If
                oWorksheet.Activate
                .Windows(1).View = 1                      'xlNormalView �W��
                .Windows(1).Zoom = 25                     '�\���{��
                .Windows(1).ScrollColumn = 1              '��1�����[�ɂȂ�悤�ɃE�B���h�E���X�N���[��
                .Windows(1).ScrollRow = 1                 '�s1����[�ɂȂ�悤�ɃE�B���h�E���X�N���[��
                .Windows(1).FreezePanes = False           '�E�B���h�E�g�̌Œ����

                '�V�[�g����ύX�A�^�u�̐F��ύX
                oWorksheet.Name = oWorkSheetRenameInfo.Item(lCnt).Item("After")
                oWorksheet.Tab.ThemeColor = oStringToThemeColor.Item(asFromToString)
                oWorksheet.Tab.TintAndShade = 0

                '�V�[�g���r���ʗp�̐V�K���[�N�u�b�N�ɃR�s�[
                Call oWorksheet.Copy(, aoWorkbookForResults.Worksheets(aoWorkbookForResults.Worksheets.Count))
            End If
        Next

        '��r�Ώۃt�@�C�������
        Call .Close(False)
    End With
    
    If Len(sTempPath) Then
    '�}�N������̏ꍇ�ɕʖ��ŕۑ������t�@�C������������폜����
        Call func_CM_DeleteFile(sTempPath)
    End If

    '�T�}���[�V�[�g�̃J�����ʒu�ϊ��p�n�b�V���}�b�v��`
    Dim oStringToColumn : Set oStringToColumn = CreateObject("Scripting.Dictionary")
    Call oStringToColumn.Add("From",1)
    Call oStringToColumn.Add("To",2)
    
    '�T�}���[�V�[�g�ɔ�r�Ώۃt�@�C���̏����o��
    Dim lRow : Dim lColumn : Dim oItem
    lColumn = oStringToColumn.Item(asFromToString)
    With aoWorkbookForResults.Worksheets.Item(1)
        '�t�@�C���p�X
        lRow = 1
        .Cells(lRow, lColumn).Value = asPath
        '�V�[�g��
        For Each oItem In oWorkSheetRenameInfo.Items
            lRow = lRow + 1
            .Cells(lRow, lColumn).Value = oItem.Item("Before")
        Next
    End With

    '���[�N�V�[�g�̃��l�[������ԋp
    Set func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = oWorkSheetRenameInfo

    '�I�u�W�F�N�g���J��
    Set oStringToColumn = Nothing
    Set oItem = Nothing
    Set oWorksheet = Nothing
    Set oStringToThemeColor = Nothing
    Set oWorkSheetRenameInfo = Nothing
    Set oWorkBook = Nothing
    Set oExcel = Nothing
    
End Function

'�B�|�Q�|�P�[�P�D�ꎞ�t�@�C���̃p�X���擾
Private Function func_CmpExcelGetTempFilePath()
    Dim sParentFolderPath : sParentFolderPath = func_CM_GetParentFolderPath(WScript.ScriptFullName)
    Dim sFolderPath : sFolderPath = func_CM_BuildPath(sParentFolderPath, Cs_FOLDER_TEMP)
    func_CmpExcelGetTempFilePath = func_CM_BuildPath(sFolderPath, func_CM_GetTempFileName())
End Function

'�B�|�Q�|�P�[�Q�D�V�[�g���쐬
Private Function func_CmpExcelMakeSheetName( _
    byVal alCnt _
    , byVal asFromToString _
    )
    func_CmpExcelMakeSheetName = "�y" & asFromToString & "_" & CStr(alCnt) & "�V�[�g�ځz"
End Function

'�B�|�Q�|�P�[�R�D�ύX�O��̃��[�N�V�[�g���i�[�p�n�b�V���}�b�v�쐬
Private Function func_CmpExcelGetMapWorkSheetRenameInfo( _
    byVal asBefore _
    , byVal asAfter _
    )
    Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
    Call oTemp.Add("Before", asBefore)
    Call oTemp.Add("After", asAfter)
    Set func_CmpExcelGetMapWorkSheetRenameInfo = oTemp
    Set oTemp = Nothing
End Function

'�B�|�R�D��r����
Private Sub sub_CmpExcelCompare( _
    byRef aoParams _
    , byRef aoWorkbookForResults _
    )
    '���[�N�V�[�g���Ƃ̃V�[�g�����l�[�����p�n�b�V���}�b�v
    Dim oWorkSheetRenameInfos : Set oWorkSheetRenameInfos = aoParams.Item("WorkSheetRenameInfos")
    Dim oFrom : Set oFrom = oWorkSheetRenameInfos.Item("From")
    Dim oTo : Set oFrom = oWorkSheetRenameInfos.Item("To")

    Dim lCnt
    For lCnt = 1 To func_CM_Min(oFrom.Count, oTo.Count)
    '��r����̊e�V�[�g�ɍ����������鏑���ݒ������
        '��r���iTo�j�̃V�[�g�ɑ΂���r��iFrom�j�Ƃ̍�����������悤�ɂ���
        Call sub_CmpExcelSetFormatToUnderstandDifference(_
                aoWorkbookForResults, oFrom.Item(lCnt).Item("After"), oTo.Item(lCnt).Item("After"))        
        '��r��iFrom�j�̃V�[�g�ɑ΂���r���iTo�j�Ƃ̍�����������悤�ɂ���
        Call sub_CmpExcelSetFormatToUnderstandDifference( _
                aoWorkbookForResults, oTo.Item(lCnt).Item("After"), oFrom.Item(lCnt).Item("After"))        
    Next

    '�I�u�W�F�N�g���J��
    Set oTo = Nothing
    Set oFrom = Nothing
    Set oWorkSheetRenameInfos = Nothing

End Sub

'�B�|�R�[�P�DasSheetNameA�̃V�[�g��asSheetNameB�V�[�g�Ƃ̍����������鏑���ݒ������
Private Sub sub_CmpExcelSetFormatToUnderstandDifference( _
    byRef aoWorkbookForResults _
    , byVal asSheetNameA _
    , byVal asSheetNameB _
    )

    '�Z���̔�r
    aoWorkbookForResults.Worksheets(asSheetNameA).Activate
    aoWorkbookForResults.Worksheets(asSheetNameA).UsedRange.Select
    Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
    Call oExcel.Selection.FormatConditions.Add( _
            2 _
            , _
            , "=EXACT(OFFSET($A$1,ROW()-1,COLUMN()-1),OFFSET('" _
            & asSheetNameB _
            & "'!$A$1,ROW()-1,COLUMN()-1))=TRUE" _
            )    'xlExpression=2�i�����j
    oExcel.Selection.FormatConditions(oExcel.Selection.FormatConditions.Count).SetFirstPriority

    With oExcel.Selection.FormatConditions(1).Interior
        .Pattern = 1                        '���H xlSolid
        .PatternColorIndex = -4105          '���� xlAutomatic
        .ThemeColor = 1                     '�Z�F xlThemeColorDark1
        .TintAndShade = -0.149998474074526  '�F�𖾂邭���邩�܂��͈Â�����
        .PatternTintAndShade = 0            '�Z�F�ƖԊ|���p�^�[��
    End With

    With oExcel.Selection.FormatConditions(1).Font
        .ThemeColor = 1                     '�Z�F xlThemeColorDark1
        .TintAndShade = -0.499984740745262  '�F�𖾂邭���邩�܂��͈Â�����
    End With

    aoWorkbookForResults.Worksheets(asSheetNameA).Range("A1").Select

    '�I�[�g�V�F�C�v�̔�r
    Dim oAutoshapeA : Dim oAutoshapeB
    For Each oAutoshapeA In aoWorkbookForResults.Worksheets(asSheetNameA).Shapes
        Set oAutoshapeB = func_CM_GetObjectByIdFromCollection(aoWorkbookForResults.Worksheets(asSheetNameA).Shapes, oAutoshapeA.Id)
        If Trim(func_CM_ExcelGetTextFromAutoshape(oAutoshapeA)) _
           = Trim(func_CM_ExcelGetTextFromAutoshape(oAutoshapeB)) Then
        '�I�[�g�V�F�C�v��ID�ƃe�L�X�g����v����i���ق��Ȃ��j�ꍇ�͊D�F�ɂ���
            Call sub_CmpExcelSetAutoshapeColor(oAutoshapeA)
        End If
    Next

    '�I�u�W�F�N�g���J��
    Set oAutoshape = Nothing
    Set oExcel = Nothing
End Sub

'�B�|�R�[�P�[�P�D�I�[�g�V�F�C�v�̐F���D�F�ɂ���
Private Sub sub_CmpExcelSetAutoshapeColor( _
    byRef aoAutoshape _
    )
    On Error Resume Next
    With aoAutoshape.Fill
        .Visible = True                          'msoTrue
        .ForeColor.ObjectTjemeColor = 14         '�w�i�P�e�[�}�̐F msoThemeColorBackground1
        .ForeColor.TintAndShade = 0              '�F�𖾂邭���邩�܂��͈Â�����P���x���������_�^ (Single) �̒l
        .ForeColor.Brightness = -0.150000006     '���x
        .Transparency = 0                        '�h��Ԃ��̓����x������ 0.0 (�s����) ���� 1.0 (����) �܂ł̒l
        .Solid                                   '�h��Ԃ����ψ�ȐF�ɐݒ�
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub
