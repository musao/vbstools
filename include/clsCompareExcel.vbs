'***************************************************************************************************
'FILENAME                    : clsCompareExcel.vbs
'Overview                    : �G�N�Z���t�@�C���̔�r���s��
'Detailed Description        : ���ʊ֐����C�u�����iVbsBasicLibCommon.vbs�j��ǂݍ���ł���g�p���邱��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCompareExcel
    '�N���X���ϐ��A�萔
    Private PsPathFrom, PsPathTo, PoPubSub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : �R���X�g���N�^
    'Detailed Description        : �����ϐ��̏�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        '������
        PsPathFrom = ""
        PsPathTo = ""
        Set PoPubSub = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : �f�X�g���N�^
    'Detailed Description        : �I������
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoPubSub = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let PathFrom()
    'Overview                    : ��r���G�N�Z���t�@�C���̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : ��r����G�N�Z���t�@�C���̃p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let PathFrom( _
        byVal asPath _
        )
        If func_CM_FsFileExists(asPath) Then PsPathFrom = asPath Else PsPathFrom = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get PathFrom()
    'Overview                    : ��r���G�N�Z���t�@�C���̃p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��r���G�N�Z���t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get PathFrom()
        PathFrom = PsPathFrom
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let PathTo()
    'Overview                    : ��r��G�N�Z���t�@�C���̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : ��r����G�N�Z���t�@�C���̃p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let PathTo( _
        byVal asPath _
        )
        If func_CM_FsFileExists(asPath) Then PsPathTo = asPath Else PsPathTo = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get PathTo()
    'Overview                    : ��r��G�N�Z���t�@�C���̃p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��r��G�N�Z���t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get PathTo()
        PathTo = PsPathTo
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set PubSub()
    'Overview                    : �o��-�w�ǌ^�iPublish/subscribe�j�N���X�̃I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoPubSub               : �o��-�w�ǌ^�iPublish/subscribe�j�N���X�̃I�u�W�F�N�g
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set PubSub( _
        byRef aoPubSub _
        )
        Set PoPubSub = aoPubSub
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get PubSub()
    'Overview                    : �o��-�w�ǌ^�iPublish/subscribe�j�N���X�̃I�u�W�F�N�g��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o��-�w�ǌ^�iPublish/subscribe�j�N���X�̃I�u�W�F�N�g
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get PubSub()
        Set PubSub = PoPubSub
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Compare()
    'Overview                    : �G�N�Z���t�@�C�����r����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True:���튮�� / False:���s
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Compare( _
        )
        Dim sMyName : sMyName = "+Compare"
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "PsPathFrom = " & func_CM_ToString(PsPathFrom) & ", PsPathTo = " & func_CM_ToString(PsPathTo))
        
        Compare = False
        
        '��r���ʗp�̐V�K���[�N�u�b�N���쐬
        With CreateObject("Excel.Application")
            .DisplayAlerts = False
            .ScreenUpdating = False
            .AutomationSecurity = 3                               'msoAutomationSecurityForceDisable = 3
            Dim oWorkbookForResults
            Set oWorkbookForResults = .Workbooks.Add(-4167)      '�V�K���[�N�u�b�N xlWBATWorksheet=-4167
        End With
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Create a new workbook for comparison.")
        
        Dim oParams : Set oParams = new_DictSetValues(Array("WorkbookForResults", oWorkbookForResults))
        
        '��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
        Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(oParams)
        
        '�G�N�Z���t�@�C�����r����
        Call sub_CmpExcelCompare(oParams)
        
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        
        '�I��
        Set oParams = Nothing
        Set oWorkbookForResults = Nothing
        Compare = True
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCopyAllSheetsToWorkbookForResults()
    'Overview                    : ��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
    'Detailed Description        : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g�Ɋi�[����
    '                              ���[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v�̍\��
    '                              Key                       Value
    '                              --------------------      -------------------------------------------
    '                              "WorkbookForResults"      ��r���ʗp�̃��[�N�u�b�N
    '                              "From"                    ��r�����[�N�V�[�g�̃��l�[�����iclsCmArray�^�j
    '                              "To"                      ��r�惏�[�N�V�[�g�̃��l�[�����iclsCmArray�^�j
    'Argument
    '     aoParams               : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCopyAllSheetsToWorkbookForResults( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelCopyAllSheetsToWorkbookForResults"
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, func_CM_ToString(aoParams))
        
        '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g����K�v�ȗv�f�����o��
        Dim oWorkbookForResults : Call sub_CM_Bind(oWorkbookForResults, aoParams.Item("WorkbookForResults"))
        
        Dim sPath : Dim sFromToString
        '��r���t�@�C���̃R�s�[
        sPath = PsPathFrom : sFromToString = "From" 
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString))
        
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Source file copy completed.")
        Call sub_CmpExcelPublish("log", 9, sMyName, func_CM_ToString(aoParams))
        
        '��r��t�@�C���̃R�s�[
        sPath = PsPathTo : sFromToString = "To"
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString))
        
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        Call sub_CmpExcelPublish("log", 9, sMyName, func_CM_ToString(aoParams))
        
        Set oWorkbookForResults = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail()
    'Overview                    : ��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
    'Detailed Description        : ��r�Ώۂ̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[������ŁA
    '                              �V�[�g���Ƃ̕ύX�O��̃V�[�g�����i�[�����I�u�W�F�N�g�i�ȉ��Q�Ɓj
    '                              �̔z��iclsCmArray�^�j��Ԃ�
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "Before"                 �ύX�O�̃��[�N�V�[�g��
    '                              "After"                  �ύX��̃��[�N�V�[�g��
    'Argument
    '     aoWorkbookForResults   : ��r���ʗp�̃��[�N�u�b�N
    '     asPath                 : ��r�Ώۃt�@�C���̃p�X
    '     asFromToString         : ��r��������ʂ��镶���� "From","To"
    'Return Value
    '     �V�[�g���Ƃ̕ύX�O��̃V�[�g�����i�[�����I�u�W�F�N�g�̔z��iclsCmArray�^�j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail( _
        byRef aoWorkbookForResults _
        , byVal asPath _
        , byVal asFromToString _
        )
        Dim sMyName : sMyName = "-func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail"
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "aoWorkbookForResults = " & func_CM_ToString(aoWorkbookForResults) & ", asPath = " & func_CM_ToString(asPath)& ", asFromToString = " & func_CM_ToString(asFromToString))

        '��r�Ώۃt�@�C�����J��
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        Dim oWorkBook : Set oWorkBook = func_CM_ExcelOpenFile(oExcel, asPath)
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Opened Excel file, file path is " & func_CM_ToString(asPath) )
        
        Dim sTempPath : sTempPath = ""
        If oWorkBook.HasVBProject Then
        '�}�N������̏ꍇ�͕ʖ��ŕۑ�������ōēx�J��
            sTempPath = func_CM_FsGetTempFilePath()
            Call sub_CM_ExcelSaveAs(oWorkBook, sTempPath, vbNullString)
            Set oWorkBook = func_CM_ExcelOpenFile( oExcel, sTempPath)
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "It was Excel with a macro, so save it with a different name and reopen it.")
        End If

        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Attempt to unprotect Excel file." )
        '�����̕ی����������
        Call sub_CM_OfficeUnprotect(oWorkBook, vbNullString)
        If Err.Number<>0 Then
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "It couldn't." )
            Call sub_CmpExcelPublish("log", 9, sMyName, func_CM_ToStringErr() )
            Err.Clear
        End If
        
        With oWorkBook
            '���[�N�V�[�g�̃��l�[�����i�[�p�z��iclsCmArray�^�j
            Dim oWorkSheetRenameInfo : Set oWorkSheetRenameInfo = new_clsCmArray()
            '�^�u�̐F�ϊ��p�n�b�V���}�b�v��`
            Dim oStringToThemeColor : Set oStringToThemeColor = new_DictSetValues(Array("From", 2, "To", 8))
            
            Dim oWorksheet, sNewSheetName
            For Each oWorksheet In .Worksheets
                If oWorksheet.Visible=True Then
                '�S�Ă̌�����V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
                    '�����O�o��
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Start processing sheet " & func_CM_ToString(oWorksheet.Name) & "." )
                    
                    '���[�N�V�[�g���擾����ѕύX���閼�̂����߂�
                    sNewSheetName = func_CmpExcelMakeSheetName(oWorkSheetRenameInfo.Length+1, asFromToString)
                    oWorkSheetRenameInfo.Push new_DictSetValues( Array("Before", oWorksheet.Name, "After", sNewSheetName) )
                    '�����O�o��
                    Call sub_CmpExcelPublish("log", 9, sMyName, "oWorkSheetRenameInfo = " & func_CM_ToString(oWorkSheetRenameInfo) )
                    
                    '�V�[�g�ی�̉���
                    Call sub_CmpExcelTryToSomething(1, oWorksheet, "Try to unprotect a sheet.")
                    
                    '�I�[�g�t�B���^�̉���
                    Call sub_CmpExcelTryToSomething(2, oWorksheet, "Try to clear the AutoFilter.")
                    
                    '�V�[�g���ύX���^�u�̐F��ύX
                    oWorksheet.Name = sNewSheetName
                    oWorksheet.Tab.ThemeColor = oStringToThemeColor.Item(asFromToString)
                    oWorksheet.Tab.TintAndShade = 0
                    '�V�[�g�̕\���𐮂���
                    oWorksheet.Activate
                    .Windows(1).View = 1                      'xlNormalView �W��
                    .Windows(1).Zoom = 25                     '�\���{��
                    .Windows(1).ScrollColumn = 1              '��1�����[�ɂȂ�悤�ɃE�B���h�E���X�N���[��
                    .Windows(1).ScrollRow = 1                 '�s1����[�ɂȂ�悤�ɃE�B���h�E���X�N���[��
                    .Windows(1).FreezePanes = False           '�E�B���h�E�g�̌Œ����
                    
                    '�����O�o��
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Start copying sheets to a new workbook for comparison results.")
                    '�V�[�g���r���ʗp�̐V�K���[�N�u�b�N�ɃR�s�[
                    Call oWorksheet.Copy(, aoWorkbookForResults.Worksheets(aoWorkbookForResults.Worksheets.Count))
                    '�����O�o��
                    Call sub_CmpExcelPublish("log", 3, sMyName, "Copy Complete.")
                End If
            Next

            '��r�Ώۃt�@�C�������
            Call .Close(False)
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "Close the file being compared." )
        End With
        
        If Len(sTempPath) Then
        '�}�N������̏ꍇ�ɕʖ��ŕۑ������t�@�C������������폜����
            Call func_CM_FsDeleteFile(sTempPath)
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "Delete file saved with a different name.")
        End If

        '�T�}���[�V�[�g�̃J�����ʒu�ϊ��p�n�b�V���}�b�v��`
        Dim oStringToColumn : Set oStringToColumn = new_DictSetValues(Array("From", 1, "To", 2))
        '�T�}���[�V�[�g�ɔ�r�Ώۃt�@�C���̏����o��
        Dim lRow : Dim lColumn : Dim oItem
        lColumn = oStringToColumn.Item(asFromToString)
        With aoWorkbookForResults.Worksheets.Item(1)
            '�t�@�C���p�X
            lRow = 1
            .Cells(lRow, lColumn).Value = asPath
            '�V�[�g��
            For Each oItem In oWorkSheetRenameInfo.Map(new_Func( "(e,i,a)=>e.Item(""Before"")" ) ).Items
                lRow = lRow + 1
                .Cells(lRow, lColumn).Value = oItem
            Next
        End With
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Output the information of the files to be compared in the summary sheet.")

        '���[�N�V�[�g�̃��l�[������ԋp
        Set func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = oWorkSheetRenameInfo
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        Call sub_CmpExcelPublish("log", 9, sMyName, "func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail = " & func_CM_ToString(oWorkSheetRenameInfo))
        
        '�I�u�W�F�N�g���J��
        Set oStringToColumn = Nothing
        Set oItem = Nothing
        Set oWorksheet = Nothing
        Set oStringToThemeColor = Nothing
        Set oWorkSheetRenameInfo = Nothing
        Set oWorkBook = Nothing
        Set oExcel = Nothing
        
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelMakeSheetName()
    'Overview                    : �V�[�g���쐬
    'Detailed Description        : �H����
    'Argument
    '     alCnt                  : �V�[�g�̐擪����̔ԍ�
    '     asFromToString         : ��r��������ʂ��镶���� "From","To"
    'Return Value
    '     �V�[�g��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelMakeSheetName( _
        byVal alCnt _
        , byVal asFromToString _
        )
        func_CmpExcelMakeSheetName = "�y" & asFromToString & "_" & CStr(alCnt) & "�V�[�g�ځz"
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCompare()
    'Overview                    : �G�N�Z���t�@�C�����r����
    'Detailed Description        : �H����
    'Argument
    '     aoParams               : �p�����[�^�i�[�p�ėp�I�u�W�F�N�g
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCompare( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelCompare"
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "Start")
        Call sub_CmpExcelPublish("log", 9, sMyName, "aoParams = " & func_CM_ToString(aoParams))
        
        '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g����K�v�ȗv�f�����o��
        Dim oWorkbookForResults : Call sub_CM_Bind(oWorkbookForResults, aoParams.Item("WorkbookForResults"))
        Dim oFrom : Call sub_CM_Bind(oFrom, aoParams.Item("From"))
        Dim oTo : Call sub_CM_Bind(oTo, aoParams.Item("To"))

        Dim lCnt
        For lCnt = 0 To func_CM_MathMin(oFrom.Length, oTo.Length)-1
        '��r����̊e�V�[�g�ɍ����������鏑���ݒ������
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "Comparison of " & lCnt+1 & "th sheets.")
            
            '��r���iTo�j�̃V�[�g�ɑ΂���r��iFrom�j�Ƃ̍�����������悤�ɂ���
            Call sub_CmpExcelSetFormatToUnderstandDifference(_
                    oWorkbookForResults, oFrom(lCnt).Item("After"), oTo(lCnt).Item("After"))
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "to see the difference from the comparison destination (" & oFrom(lCnt).Item("Before") & ") to the source sheet (" & oTo(lCnt).Item("Before") & ").")
            
            '��r��iFrom�j�̃V�[�g�ɑ΂���r���iTo�j�Ƃ̍�����������悤�ɂ���
            Call sub_CmpExcelSetFormatToUnderstandDifference( _
                    oWorkbookForResults, oTo(lCnt).Item("After"), oFrom(lCnt).Item("After"))
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, "to see the difference from the comparison source (" & oTo(lCnt).Item("Before") & ") to the comparison destination sheet (" & oFrom(lCnt).Item("Before") & ").")
            
        Next
        
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Arrange the Window so that you can see the difference.")
        '�����u�b�N�̐V�����E�B���h�E���J��
        oWorkbookForResults.Worksheets(oFrom(0).Item("After")).Activate
        With oWorkbookForResults.Windows(1).NewWindow
            Dim sCaption : sCaption = .Caption
            Dim oWorksheet
            For Each oWorksheet In .Parent.Worksheets
                oWorksheet.Activate
                .Zoom = 25
            Next
        End With
        oWorkbookForResults.Worksheets(oTo(0).Item("After")).Activate
        '���ׂĔ�r
        oWorkbookForResults.Activate
        With oWorkbookForResults.Parent
            .Windows.CompareSideBySideWith(sCaption)
            Call .Windows.Arrange(-4166, True)               'xlVertical = -4166
            .DisplayAlerts = True
            .ScreenUpdating = True
            .AutomationSecurity = 2                     'msoAutomationSecurityByUI = 2 [ �Z�L�����e�B] �_�C�A���O �{�b�N�X�Ŏw�肳�ꂽ�Z�L�����e�B�ݒ���g�p
            .Visible = True
        End With
        
        '�����O�o��
        Call sub_CmpExcelPublish("log", 5, sMyName, "End")
        
        '�I�u�W�F�N�g���J��
        Set oWorksheet = Nothing
        Set oTo = Nothing
        Set oFrom = Nothing
        Set oWorkbookForResults = Nothing

    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelSetFormatToUnderstandDifference()
    'Overview                    : asSheetNameA�̃V�[�g��asSheetNameB�V�[�g�Ƃ̍����������鏑���ݒ������
    'Detailed Description        : �H����
    'Argument
    '     aoWorkbookForResults   : ��r���ʗp�̃��[�N�u�b�N
    '     asSheetNameA           : ��r���̃V�[�g��
    '     asSheetNameB           : ��r��̃V�[�g��
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelSetFormatToUnderstandDifference( _
        byRef aoWorkbookForResults _
        , byVal asSheetNameA _
        , byVal asSheetNameB _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelSetFormatToUnderstandDifference"

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
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "Cell comparison complete.")

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
        '�����O�o��
        Call sub_CmpExcelPublish("log", 3, sMyName, "AutoShape comparison complete.")

        '�I�u�W�F�N�g���J��
        Set oAutoshapeA = Nothing
        Set oAutoshapeB = Nothing
        Set oExcel = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelSetAutoshapeColor()
    'Overview                    : �I�[�g�V�F�C�v�̐F���D�F�ɂ���
    'Detailed Description        : �G���[�͖�������
    'Argument
    '     aoAutoshape            : �I�[�g�V�F�C�v
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
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

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelTryToSomething()
    'Overview                    : �V�[�g�̐ݒ�ύX�����݂�
    'Detailed Description        : �H����
    'Argument
    '     alType                 : ���
    '                               1:�V�[�g���ی삳��Ă������������
    '                               2:�I�[�g�t�B���^���ݒ肳��Ă������������
    '     aoWorksheet            : ���[�N�V�[�g
    '     asLogMsg               : ���O�ɏo�͂��郁�b�Z�[�W
    'Return Value
    '     �V�[�g��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelTryToSomething( _
        byVal alType _
        , byRef aoWorksheet _
        , byVal asLogMsg _
        )
        Dim sMyName : sMyName = "-sub_CmpExcelTryToSomething"
        
        '�m�F���e
        Dim boFlg
        Select Case alType
            Case 1:
                boFlg = aoWorksheet.ProtectContents
            Case 2:
                boFlg = aoWorksheet.AutoFilterMode
        End Select
        
        If boFlg Then
            '�����O�o��
            Call sub_CmpExcelPublish("log", 3, sMyName, asLogMsg)
            
            On Error Resume Next
            
            '�ݒ�ύX�����݂�
            Select Case alType
                Case 1:
                    aoWorksheet.Unprotect
                Case 2:
                    aoWorksheet.Cells(1,1).AutoFilter
            End Select
            
            '����
            If Err.Number<>0 Then
                '�����O�o��
                Call sub_CmpExcelPublish("log", 3, sMyName, "It couldn't." )
                Call sub_CmpExcelPublish("log", 9, sMyName, func_CM_ToStringErr() )
                Err.Clear
            End If
        End If
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelPublish()
    'Overview                    : �o�ŁiPublish�j����
    'Detailed Description        : �o��-�w�ǌ^�iPublish/subscribe�j�N���X������Ώo�ŁiPublish�j��������
    'Argument
    '     asTopic                : �g�s�b�N
    '     alLevel                : ���x��
    '     asFuncName             : �֐���
    '     asCont                 : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelPublish( _
        byVal asTopic _
        , byVal alLevel _
        , byVal asFuncName _
        , byVal asCont _
        )
        If PoPubSub Is Nothing Then Exit Sub
        Call PoPubSub.Publish(asTopic, Array(alLevel, TypeName(Me)&asFuncName, asCont))
    End Sub
    
End Class
