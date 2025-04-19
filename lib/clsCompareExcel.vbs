'***************************************************************************************************
'FILENAME                    : clsCompareExcel.vbs
'Overview                    : �G�N�Z���t�@�C���̔�r���s��
'Detailed Description        : ���ʊ֐����C�u������ǂݍ���ł���g�p���邱��
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCompareExcel
    '�N���X���ϐ��A�萔
    Private PsPathFrom, PsPathTo, PoBroker
    
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
        Set PoBroker = Nothing
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
        Set PoBroker = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let pathFrom()
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
    Public Property Let pathFrom( _
        byVal asPath _
        )
        If new_Fso().FileExists(asPath) Then PsPathFrom = asPath Else PsPathFrom = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get pathFrom()
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
    Public Property Get pathFrom()
        pathFrom = PsPathFrom
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let pathTo()
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
    Public Property Let pathTo( _
        byVal asPath _
        )
        If new_Fso().FileExists(asPath) Then PsPathTo = asPath Else PsPathTo = ""
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get pathTo()
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
    Public Property Get pathTo()
        pathTo = PsPathTo
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set broker()
    'Overview                    : �u���[�J�[�N���X�̃I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoBroker               : �u���[�J�[�N���X�̃C���X�^���X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set broker( _
        byRef aoBroker _
        )
        Set PoBroker = aoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get broker()
    'Overview                    : �u���[�J�[�N���X�̃I�u�W�F�N�g��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �u���[�J�[�N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get broker()
        Set broker = PoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : compare()
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
    Public Function compare( _
        )
        Dim sMyName : sMyName = "+compare()"
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "Start"
        this_publishLog logType.TRACE, sMyName, "PsPathFrom = " & cf_toString(PsPathFrom) & ", PsPathTo = " & cf_toString(PsPathTo)
        
        compare = False
        
        '��r���ʗp�̐V�K���[�N�u�b�N���쐬
        With CreateObject("Excel.Application")
            .DisplayAlerts = False
            .ScreenUpdating = False
            .AutomationSecurity = 3                               'msoAutomationSecurityForceDisable = 3
            Dim oWorkbookForResults
            Set oWorkbookForResults = .Workbooks.Add(-4167)      '�V�K���[�N�u�b�N xlWBATWorksheet=-4167
        End With
        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Create a new workbook for comparison."
        
        Dim oParams : Set oParams = new_DicOf(Array("WorkbookForResults", oWorkbookForResults))
        
        '��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
        this_copyAllSheetsToWorkbookForResults oParams
        
        '�G�N�Z���t�@�C�����r����
        this_compare oParams
        
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "End"
        
        '�I��
        Set oParams = Nothing
        Set oWorkbookForResults = Nothing
        compare = True
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_copyAllSheetsToWorkbookForResults()
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
    Private Sub this_copyAllSheetsToWorkbookForResults( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-this_copyAllSheetsToWorkbookForResults()"
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "Start"
        this_publishLog logType.TRACE, sMyName, cf_toString(aoParams)
        
        '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g����K�v�ȗv�f�����o��
        Dim oWorkbookForResults : cf_bind oWorkbookForResults, aoParams.Item("WorkbookForResults")
        
        Dim sPath : Dim sFromToString
        '��r���t�@�C���̃R�s�[
        sPath = PsPathFrom : sFromToString = "From" 
        aoParams.Add sFromToString, _
            this_copyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString)
        
        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Source file copy completed."
        this_publishLog logType.TRACE, sMyName, cf_toString(aoParams)
        
        '��r��t�@�C���̃R�s�[
        sPath = PsPathTo : sFromToString = "To"
        aoParams.Add sFromToString, _
            this_copyAllSheetsToWorkbookForResultsDetail(oWorkbookForResults, sPath, sFromToString)
        
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "End"
        this_publishLog logType.TRACE, sMyName, cf_toString(aoParams)
        
        Set oWorkbookForResults = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_copyAllSheetsToWorkbookForResultsDetail()
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
    Private Function this_copyAllSheetsToWorkbookForResultsDetail( _
        byRef aoWorkbookForResults _
        , byVal asPath _
        , byVal asFromToString _
        )
        Dim sMyName : sMyName = "-this_copyAllSheetsToWorkbookForResultsDetail()"
        
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "Start"
        this_publishLog logType.TRACE, sMyName, "aoWorkbookForResults = " & cf_toString(aoWorkbookForResults) & ", asPath = " & cf_toString(asPath)& ", asFromToString = " & cf_toString(asFromToString)

        '��r�Ώۃt�@�C�����J��
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        Dim oWorkBook : Set oWorkBook = func_CM_ExcelOpenFile(oExcel, asPath)
        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Opened Excel file, file path is " & cf_toString(asPath)
        
        Dim sTempPath : sTempPath = ""
        If oWorkBook.HasVBProject Then
        '�}�N������̏ꍇ�͕ʖ��ŕۑ�������ōēx�J��
            sTempPath = fw_getTempPath()
            sub_CM_ExcelSaveAs oWorkBook, sTempPath, vbNullString
            Set oWorkBook = func_CM_ExcelOpenFile( oExcel, sTempPath)
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "It was Excel with a macro, so save it with a different name and reopen it."
        End If

        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Attempt to unprotect Excel file."
        '�����̕ی����������
        this_tryCatchAfterProc fw_tryCatch(new_Func("a=>a.Unprotect"), oWorkBook, empty, empty), sMyName
        
        With oWorkBook
            '���[�N�V�[�g�̃��l�[�����i�[�p�z��iclsCmArray�^�j
            Dim oWorkSheetRenameInfo : Set oWorkSheetRenameInfo = new_Arr()
            '�^�u�̐F�ϊ��p�n�b�V���}�b�v��`
            Dim oStringToThemeColor : Set oStringToThemeColor = new_DicOf(Array("From", 2, "To", 8))
            
            Dim oWorksheet, sNewSheetName
            For Each oWorksheet In .Worksheets
                If oWorksheet.Visible=True Then
                '�S�Ă̌�����V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
                    '�����O�o��
                    this_publishLog logType.WARNING, sMyName, "Start processing sheet " & cf_toString(oWorksheet.Name) & "."
                    
                    '�V�[�g�ی�̉���
                    this_publishLog logType.WARNING, sMyName, "Try to unprotect a sheet."
                    this_tryCatchAfterProc fw_tryCatch(new_Func("a=>{If a.ProtectContents Then:a.Unprotect(vbNullString):End If}"), oWorksheet, empty, empty), sMyName
                    
                    '�I�[�g�t�B���^�̉���
                    this_publishLog logType.WARNING, sMyName, "Try to clear the AutoFilter."
                    this_tryCatchAfterProc fw_tryCatch(new_Func("a=>{If a.AutoFilterMode Then:a.Cells(1,1).AutoFilter:End If}"), oWorksheet, empty, empty), sMyName
                    
                    '���[�N�V�[�g���擾����ѕύX���閼�̂����߂�
                    sNewSheetName = this_makeSheetName(oWorkSheetRenameInfo.Length+1, asFromToString)
                    oWorkSheetRenameInfo.Push new_DicOf( Array("Before", oWorksheet.Name, "After", sNewSheetName) )
                    '�����O�o��
                    this_publishLog logType.TRACE, sMyName, "oWorkSheetRenameInfo = " & cf_toString(oWorkSheetRenameInfo)
                    
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
                    this_publishLog logType.WARNING, sMyName, "Start copying sheets to a new workbook for comparison results."
                    '�V�[�g���r���ʗp�̐V�K���[�N�u�b�N�ɃR�s�[
                    oWorksheet.Copy , aoWorkbookForResults.Worksheets(aoWorkbookForResults.Worksheets.Count)
                    '�����O�o��
                    this_publishLog logType.WARNING, sMyName, "Copy Complete."
                End If
            Next

            '��r�Ώۃt�@�C�������
            .Close False
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "Close the file being compared."
        End With
        
        If Len(sTempPath) Then
        '�}�N������̏ꍇ�ɕʖ��ŕۑ������t�@�C������������폜����
            fs_deleteFile sTempPath
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "Delete file saved with a different name."
        End If

        '�T�}���[�V�[�g�̃J�����ʒu�ϊ��p�n�b�V���}�b�v��`
        Dim oStringToColumn : Set oStringToColumn = new_DicOf(Array("From", 1, "To", 2))
        '�T�}���[�V�[�g�ɔ�r�Ώۃt�@�C���̏����o��
        Dim lRow : Dim lColumn : Dim oItem
        lColumn = oStringToColumn.Item(asFromToString)
        With aoWorkbookForResults.Worksheets.Item(1)
            '�t�@�C���p�X
            lRow = 1
            .Cells(lRow, lColumn).Value = asPath
            '�V�[�g��
            For Each oItem In oWorkSheetRenameInfo.map(new_Func( "(e,i,a)=>e.Item(""Before"")" ) ).items
                lRow = lRow + 1
                .Cells(lRow, lColumn).Value = oItem
            Next
        End With
        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Output the information of the files to be compared in the summary sheet."

        '���[�N�V�[�g�̃��l�[������ԋp
        Set this_copyAllSheetsToWorkbookForResultsDetail = oWorkSheetRenameInfo
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "End"
        this_publishLog logType.TRACE, sMyName, "this_copyAllSheetsToWorkbookForResultsDetail = " & cf_toString(oWorkSheetRenameInfo)
        
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
    'Function/Sub Name           : this_makeSheetName()
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
    Private Function this_makeSheetName( _
        byVal alCnt _
        , byVal asFromToString _
        )
        this_makeSheetName = "�y" & asFromToString & "_" & CStr(alCnt) & "�V�[�g�ځz"
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_compare()
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
    Private Sub this_compare( _
        byRef aoParams _
        )
        Dim sMyName : sMyName = "-this_compare()"
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "Start"
        this_publishLog logType.TRACE, sMyName, "aoParams = " & cf_toString(aoParams)
        
        '�p�����[�^�i�[�p�ėp�I�u�W�F�N�g����K�v�ȗv�f�����o��
        Dim oWorkbookForResults : cf_bind oWorkbookForResults, aoParams.Item("WorkbookForResults")
        Dim oFrom : cf_bind oFrom, aoParams.Item("From")
        Dim oTo : cf_bind oTo, aoParams.Item("To")

        Dim lCnt
        For lCnt = 0 To math_min(oFrom.Length, oTo.Length)-1
        '��r����̊e�V�[�g�ɍ����������鏑���ݒ������
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "Comparison of " & lCnt+1 & "th sheets."
            
            '��r���iTo�j�̃V�[�g�ɑ΂���r��iFrom�j�Ƃ̍�����������悤�ɂ���
            this_setFormatToUnderstandDifference _
                    oWorkbookForResults, oFrom(lCnt).Item("After"), oTo(lCnt).Item("After")
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "to see the difference from the comparison destination (" & oFrom(lCnt).Item("Before") & ") to the source sheet (" & oTo(lCnt).Item("Before") & ")."
            
            '��r��iFrom�j�̃V�[�g�ɑ΂���r���iTo�j�Ƃ̍�����������悤�ɂ���
            this_setFormatToUnderstandDifference _
                    oWorkbookForResults, oTo(lCnt).Item("After"), oFrom(lCnt).Item("After")
            '�����O�o��
            this_publishLog logType.WARNING, sMyName, "to see the difference from the comparison source (" & oTo(lCnt).Item("Before") & ") to the comparison destination sheet (" & oFrom(lCnt).Item("Before") & ")."
            
        Next
        
        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "Arrange the Window so that you can see the difference."
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
            .Windows.Arrange -4166, True               'xlVertical = -4166
            .DisplayAlerts = True
            .ScreenUpdating = True
            .AutomationSecurity = 2                     'msoAutomationSecurityByUI = 2 [ �Z�L�����e�B] �_�C�A���O �{�b�N�X�Ŏw�肳�ꂽ�Z�L�����e�B�ݒ���g�p
            .Visible = True
        End With
        
        '�����O�o��
        this_publishLog logType.INFO, sMyName, "End"
        
        '�I�u�W�F�N�g���J��
        Set oWorksheet = Nothing
        Set oTo = Nothing
        Set oFrom = Nothing
        Set oWorkbookForResults = Nothing

    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setFormatToUnderstandDifference()
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
    Private Sub this_setFormatToUnderstandDifference( _
        byRef aoWorkbookForResults _
        , byVal asSheetNameA _
        , byVal asSheetNameB _
        )
        Dim sMyName : sMyName = "-this_setFormatToUnderstandDifference()"

        '�Z���̔�r
        aoWorkbookForResults.Worksheets(asSheetNameA).Activate
        aoWorkbookForResults.Worksheets(asSheetNameA).UsedRange.Select
        Dim oExcel : Set oExcel = aoWorkbookForResults.Parent
        oExcel.Selection.FormatConditions.Add _
                2 _
                , _
                , "=EXACT(OFFSET($A$1,ROW()-1,COLUMN()-1),OFFSET('" _
                & asSheetNameB _
                & "'!$A$1,ROW()-1,COLUMN()-1))=TRUE" _
                    'xlExpression=2�i�����j
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
        this_publishLog logType.WARNING, sMyName, "Cell comparison complete."

        '�I�[�g�V�F�C�v�̔�r
        Dim oAutoshapeA, oAutoshapeB, oRet, sTextA
        For Each oAutoshapeA In aoWorkbookForResults.Worksheets(asSheetNameA).Shapes
            Set oRet = fw_tryCatch(new_Func("(a)=>a(0).Item(a(1))"), Array(aoWorkbookForResults.Worksheets(asSheetNameB).Shapes, oAutoshapeA.Name), Empty, Empty)
            If Not oRet.isErr() Then
                Set oAutoshapeB = oRet.returnValue
                Set oRet = fw_tryCatch(Getref("func_CM_ExcelGetTextFromAutoshape"), oAutoshapeA, Empty, Empty)
                If Not oRet.isErr() Then
                    sTextA = oRet.returnValue
                    Set oRet = fw_tryCatch(Getref("func_CM_ExcelGetTextFromAutoshape"), oAutoshapeB, Empty, Empty)
                End If
                If Not oRet.isErr() Then
                    If cf_isSame(sTextA, oRet.returnValue) Then
                    '�I�[�g�V�F�C�v�̖��O�ƃe�L�X�g����v����i���ق��Ȃ��j�ꍇ�͊D�F�ɂ���
                        this_setAutoshapeColor oAutoshapeA
                    End If
                End If
            End If
        Next

        '�����O�o��
        this_publishLog logType.WARNING, sMyName, "AutoShape comparison complete."

        '�I�u�W�F�N�g���J��
        Set oRet = Nothing
        Set oAutoshapeA = Nothing
        Set oAutoshapeB = Nothing
        Set oExcel = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setAutoshapeColor()
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
    Private Sub this_setAutoshapeColor( _
        byRef aoAutoshape _
        )
        On Error Resume Next
        With aoAutoshape.Fill
            .Visible = True                          'msoTrue
            .ForeColor.ObjectThemeColor = 14         '�w�i�P�e�[�}�̐F msoThemeColorBackground1
            .ForeColor.TintAndShade = 0              '�F�𖾂邭���邩�܂��͈Â�����P���x���������_�^ (Single) �̒l
            .ForeColor.Brightness = -0.150000006     '���x
            .Transparency = 0                        '�h��Ԃ��̓����x������ 0.0 (�s����) ���� 1.0 (����) �܂ł̒l
            .Solid                                   '�h��Ԃ����ψ�ȐF�ɐݒ�
        End With
        On Error Goto 0
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_tryCatchAfterProc()
    'Overview                    : TryCatch�ŃG���[���̏���
    'Detailed Description        : �H����
    'Argument
    '     aoRet                  : fw_tryCatch()�̖߂�l
    '     asYourName             : ���������s�����֐���
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_tryCatchAfterProc( _
        byRef aoRet _
        , byVal asYourName _
        )
        If Not aoRet.isErr() Then Exit Sub
        this_publishLog logType.WARNING, asYourName, "It couldn't."
        this_publishLog logType.TRACE, asYourName, "<Err> " & cf_toString(aoRet.Item("Err"))
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_publishLog()
    'Overview                    : �o�ŁiPublish�j����
    'Detailed Description        : �u���[�J�[�N���X�������LOG�����o�ŁiPublish�j����
    'Argument
    '     alType                 : �^�C�v
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
    Private Sub this_publishLog( _
        byRef alType _
        , byVal asFuncName _
        , byVal asCont _
        )
        If PoBroker Is Nothing Then Exit Sub
        PoBroker.Publish topic.LOG, Array(alType, TypeName(Me)&asFuncName, asCont)
    End Sub
    
End Class
