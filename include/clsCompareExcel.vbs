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
    Private PdtNow
    Private PdtDate
    Private PdtStart
    Private PdtEnd
    Private PsPathFrom
    Private PsPathTo
    Private Cs_FOLDER_TEMP
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '������
        PasPathA = ""
        PasPathB = ""
        Cs_FOLDER_TEMP = "tmp"
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
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
    'Function/Sub Name           : Property Get ProcDate()
    'Overview                    : �������{������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ��������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ProcDate()
        ProcDate = PdtNow
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get StartTime()
    'Overview                    : �����J�n������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �����J�n����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get StartTime()
        StartTime = func_CM_GetDateInMilliseconds(PdtDate, PdtStart)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get EndTime()
    'Overview                    : �����I��������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �����̏I������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get EndTime()
        EndTime = func_CM_GetDateInMilliseconds(PdtDate, PdtEnd)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ElapsedTime()
    'Overview                    : �����ɂ����������Ԃ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �����ɂ�����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ElapsedTime()
       ElapsedTime = PdtEnd - PdtStart
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
        Compare = False
        
        '�J�n�����̎擾
        PdtNow = Now
        PdtDate = Date
        PdtStart = Timer
        
        '��r���ʗp�̐V�K���[�N�u�b�N���쐬
        With CreateObject("Excel.Application")
            .DisplayAlerts = False
            .ScreenUpdating = False
            .AutomationSecurity = 3                               'msoAutomationSecurityForceDisable = 3
            Dim oWorkbookForResults
            Set oWorkbookForResults = .Workbooks.Add(-4167)      '�V�K���[�N�u�b�N xlWBATWorksheet=-4167
        End With
        
        Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
        
        '��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
        Call sub_CmpExcelCopyAllSheetsToWorkbookForResults(oWorkbookForResults, oParams)
        
        '�G�N�Z���t�@�C�����r����
        Call sub_CmpExcelCompare(oWorkbookForResults, oParams)
        
        '�I��
        Set oParams = Nothing
        Set oWorkbookForResults = Nothing
        PdtEnd = Timer
        Compare = True
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCopyAllSheetsToWorkbookForResults()
    'Overview                    : ��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
    'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v�Ɋi�[����
    '                              ���[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v�̍\��
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "From"                   ��r���̃��[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v
    '                              "To"                     ��r��̃��[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v
    'Argument
    '     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    '     aoWorkbookForResults   : ��r���ʗp�̃��[�N�u�b�N
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCopyAllSheetsToWorkbookForResults( _
        byRef aoWorkbookForResults _
        , byRef aoParams _
        )
        
        Dim sPath : Dim sFromToString
        '��r���t�@�C���̃R�s�[
        sPath = PsPathFrom : sFromToString = "From" 
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

        '��r��t�@�C���̃R�s�[
        sPath = PsPathTo : sFromToString = "To"
        Call aoParams.Add(sFromToString, _
            func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail(aoWorkbookForResults, sPath, sFromToString))

    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmpExcelCopyAllSheetsToWorkbookForResultsDetail()
    'Overview                    : ��r�Ώۃt�@�C���̑S�V�[�g���r���ʗp���[�N�u�b�N�ɃR�s�[����
    'Detailed Description        : ���[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v�̍\��
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              Seq(1,2,3�c)              �ύX�O��̃��[�N�V�[�g���i�[�p�n�b�V���}�b�v
    'Argument
    '     aoWorkbookForResults   : ��r���ʗp�̃��[�N�u�b�N
    '     asPath                 : ��r�Ώۃt�@�C���̃p�X
    '     asFromToString         : ��r��������ʂ��镶���� "From","To"
    'Return Value
    '     ���[�N�V�[�g�̃��l�[�����̃n�b�V���}�b�v
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
            Call func_CM_FsDeleteFile(sTempPath)
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
    'Function/Sub Name           : func_CmpExcelGetMapWorkSheetRenameInfo()
    'Overview                    : �ύX�O��̃��[�N�V�[�g���i�[�p�n�b�V���}�b�v�쐬
    'Detailed Description        : �ύX�O��̃��[�N�V�[�g���i�[�p�n�b�V���}�b�v�̍\��
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "Before"                 �ύX�O�̃V�[�g��
    '                              "After"                  �ύX��̃V�[�g��
    'Argument
    '     asBefore               : �ύX�O�̃V�[�g��
    '     asAfter                : �ύX��̃V�[�g��
    'Return Value
    '     �ύX�O��̃��[�N�V�[�g���i�[�p�n�b�V���}�b�v
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
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

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmpExcelCompare()
    'Overview                    : �G�N�Z���t�@�C�����r����
    'Detailed Description        : �H����
    'Argument
    '     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    '     aoWorkbookForResults   : ��r���ʗp�̃��[�N�u�b�N
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmpExcelCompare( _
        byRef aoWorkbookForResults _
        , byRef aoParams _
        )
        '���[�N�V�[�g���Ƃ̃V�[�g�����l�[�����p�n�b�V���}�b�v
        Dim oFrom : Set oFrom = aoParams.Item("From")
        Dim oTo : Set oTo = aoParams.Item("To")

        Dim lCnt
        For lCnt = 1 To func_CM_MathMin(oFrom.Count, oTo.Count)
        '��r����̊e�V�[�g�ɍ����������鏑���ݒ������
            '��r���iTo�j�̃V�[�g�ɑ΂���r��iFrom�j�Ƃ̍�����������悤�ɂ���
            Call sub_CmpExcelSetFormatToUnderstandDifference(_
                    aoWorkbookForResults, oFrom.Item(lCnt).Item("After"), oTo.Item(lCnt).Item("After"))        
            '��r��iFrom�j�̃V�[�g�ɑ΂���r���iTo�j�Ƃ̍�����������悤�ɂ���
            Call sub_CmpExcelSetFormatToUnderstandDifference( _
                    aoWorkbookForResults, oTo.Item(lCnt).Item("After"), oFrom.Item(lCnt).Item("After"))        
        Next

        '�����u�b�N�̐V�����E�B���h�E���J��
        aoWorkbookForResults.Worksheets(oFrom.Item(1).Item("After")).Activate
        With aoWorkbookForResults.Windows(1).NewWindow
            Dim sCaption : sCaption = .Caption
            Dim oWorksheet
            For Each oWorksheet In .Parent.Worksheets
                oWorksheet.Activate
                .Zoom = 25
            Next
        End With
        aoWorkbookForResults.Worksheets(oTo.Item(1).Item("After")).Activate
        '���ׂĔ�r
        aoWorkbookForResults.Activate
        With aoWorkbookForResults.Parent
            .Windows.CompareSideBySideWith(sCaption)
            Call .Windows.Arrange(-4166, True)               'xlVertical = -4166
            .DisplayAlerts = True
            .ScreenUpdating = True
            .AutomationSecurity = 2                     'msoAutomationSecurityByUI = 2 [ �Z�L�����e�B] �_�C�A���O �{�b�N�X�Ŏw�肳�ꂽ�Z�L�����e�B�ݒ���g�p
            .Visible = True
        End With

        '�I�u�W�F�N�g���J��
        Set oWorksheet = Nothing
        Set oTo = Nothing
        Set oFrom = Nothing

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
    'Function/Sub Name           : func_CmpExcelGetTempFilePath()
    'Overview                    : �ꎞ�t�@�C���̃t���p�X���擾
    'Detailed Description        : ���s���̃X�N���v�g�t�@�C��������t�H���_�̉��ɂ���
    '                              Cs_FOLDER_TEMP�ȉ��̈ꎞ�t�@�C���̃p�X��Ԃ�
    '                              Cs_FOLDER_TEMP�t�H���_���Ȃ��ꍇ�͍쐬����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ꎞ�t�@�C���̃t���p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2017/04/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmpExcelGetTempFilePath()
        Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
        Dim sFolderPath : sFolderPath = func_CM_FsBuildPath(sParentFolderPath, Cs_FOLDER_TEMP)
        If Not(func_CM_FsFolderExists(sFolderPath)) Then func_CM_FsCreateFolder(sFolderPath)
        func_CmpExcelGetTempFilePath = func_CM_FsBuildPath(sFolderPath, func_CM_FsGetTempFileName())
    End Function

End Class
