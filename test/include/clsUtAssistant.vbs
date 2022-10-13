'***************************************************************************************************
'FILENAME                    : clsUtAssistant.vbs
'Overview                    : �P�̃e�X�g�p�A�V�X�^���g�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Class clsUtAssistant
'    '�N���X���ϐ�
    Private PdtNow
    Private PdtDate
    Private PdtStart
    Private PdtEnd
    Private PoRecDetail
    Private PoRecDetailTitles
    Private PoRecSumTitles
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '�J�n�����̎擾
        PdtNow = Now
        PdtDate = Date
        PdtStart = Timer
        '���ʃT�}���[�̃^�C�g����`
        Set PoRecSumTitles = CreateObject("Scripting.Dictionary")
        With PoRecSumTitles
            Call .Add(1, "Result")
            Call .Add(2, "CaseCount")
            Call .Add(3, "OkCaseCount")
            Call .Add(4, "NgCaseCount")
            Call .Add(5, "Start")
            Call .Add(6, "End")
            Call .Add(7, "ElapsedTime")
        End With
        '���ʊi�[�p�n�b�V���}�b�v
        Set PoRecDetail = CreateObject("Scripting.Dictionary")
        '���ʏڍ׃n�b�V���}�b�v�Ɋi�[������̃^�C�g����`
        Set PoRecDetailTitles = CreateObject("Scripting.Dictionary")
        With PoRecDetailTitles
            Call .Add(1, "Seq")
            Call .Add(2, "CaseName")
            Call .Add(3, "Result")
            Call .Add(4, "Start")
            Call .Add(5, "End")
            Call .Add(6, "ElapsedTime")
        End With
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoRecDetailTitles = Nothing
        Set PoRecDetail = Nothing
        Set PoRecSumTitles = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCount()
    'Overview                    : ���{�����P�̃e�X�g�P�[�X����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���{�����S�P�[�X��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCount()
        CaseCount = PoRecDetail.Count
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCountOk()
    'Overview                    : ���{�����P�̃e�X�g�P�[�X�̂���������������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���{�����P�̃e�X�g�P�[�X�̂�������������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCountOk()
        CaseCountOk = func_CountCaseAs(True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCountNg()
    'Overview                    : ���{�����P�̃e�X�g�P�[�X�̂������s��������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���{�����P�̃e�X�g�P�[�X�̂������s������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCountNg()
        CaseCountNg = func_CountCaseAs(False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get StartTime()
    'Overview                    : �P�̃e�X�g�̊J�n������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �P�̃e�X�g�̊J�n����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get StartTime()
        StartTime = func_GetDateInMilliseconds(PdtDate, PdtStart)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ProcDate()
    'Overview                    : �P�̃e�X�g�̎��{������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �P�̃e�X�g�̎��{����
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
    'Function/Sub Name           : Property Get EndTime()
    'Overview                    : �P�̃e�X�g�̏I��������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ō�̒P�̃e�X�g�̏I������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get EndTime()
        EndTime = func_GetDateInMilliseconds(PdtDate, PdtEnd)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ElapsedTime()
    'Overview                    : �P�̃e�X�g���{�ɂ����������Ԃ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �P�̃e�X�g���{�ɂ�����������
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
    'Function/Sub Name           : Run()
    'Overview                    : �e�X�g���{
    'Detailed Description        : ���ʊi�[�p�n�b�V���}�b�v�̍\��
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              Seq(1,2,3�c)              ���ʏڍ׃n�b�V���}�b�v
    '                              
    '                              ���ʏڍ׃n�b�V���}�b�v�̍\��
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "CaseName"                ���s����P�[�X���i�֐����j
    '                              "Result"                  ���� True,Flase
    '                              "Start"                   �J�n����
    '                              "End"                     �I������
    'Argument
    '     asCaseName             : ���s����P�[�X���i�֐����j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Run( _
        byVal asCaseName _
        )
        '���{
        Dim dtDate : dtDate = Date()
        Dim dtStart : dtStart = Timer
        On Error Resume Next
        Dim boResult : boResult = GetRef(asCaseName)
        Dim dtEnd : dtEnd = Timer
        If Err.Number Or Not(boResult) Then boResult = False
        
        '���ʂ��L�^
        Dim lSeq : lSeq = PoRecDetail.Count+1
        Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
        With PoRecDetailTitles
            Call oTemp.Add(.Item(1), lSeq)
            Call oTemp.Add(.Item(2), asCaseName)
            Call oTemp.Add(.Item(3), boResult)
            Call oTemp.Add(.Item(4), func_GetDateInMilliseconds(dtDate, dtStart))
            Call oTemp.Add(.Item(5), func_GetDateInMilliseconds(dtDate, dtEnd))
            Call oTemp.Add(.Item(6), dtEnd-dtStart)
        End With
        Call PoRecDetail.Add(lSeq, oTemp)
        
        '�I�����Ԃ̎擾
        PdtEnd = dtEnd
        
        Set oTemp = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : OutputReportInTsvFormat()
    'Overview                    : ���ʂ�Tsv�`���ŏo�͂���
    'Detailed Description        : �H����
    'Argument
    '     asCaseName             : ���s����P�[�X���i�֐����j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function OutputReportInTsvFormat( _
        )
        
        Dim sDelimiter : sDelimiter = vbTab
        Dim sLineFeedCode : sLineFeedCode = vbCrLf
        
        '�T�}���[��
        Dim sSum : sSum = ""
        With PoRecSumTitles
            sSum = sSum & .Item(1) & sDelimiter & isAllOk & sLineFeedCode
            sSum = sSum & .Item(2) & sDelimiter & CaseCount & sLineFeedCode
            sSum = sSum & .Item(3) & sDelimiter & CaseCountOk & sLineFeedCode
            sSum = sSum & .Item(4) & sDelimiter & CaseCountNg & sLineFeedCode
            sSum = sSum & .Item(5) & sDelimiter & StartTime & sLineFeedCode
            sSum = sSum & .Item(6) & sDelimiter & EndTime & sLineFeedCode
            sSum = sSum & .Item(7) & sDelimiter & ElapsedTime & sLineFeedCode
        End With
        
        
        '�ڍו�
        Dim lKeyT : Dim lKeyC
        
        '�w�b�_�̕ҏW
        Dim sHeader : sHeader = ""
        For lKeyT=1 To PoRecDetailTitles.Count
        '�w�b�_�̓^�C�g����`�̓��e�����ɏo�͂���
            If Len(sHeader) Then sHeader = sHeader & sDelimiter
            sHeader = sHeader & PoRecDetailTitles.Item(lKeyT)
        Next
        
        '���e�̕ҏW
        Dim sContLine
        Dim sCont : sCont = ""
        For lKeyC=1 To PoRecDetail.Count
        '���e�͌��ʊi�[�p�n�b�V���}�b�v�����ɏ�������
            sContLine = ""
            For lKeyT=1 To PoRecDetailTitles.Count
            '���ʂ��ƂɃ^�C�g�����L�[�ɒl�����o��
                If Len(sContLine) Then sContLine = sContLine & sDelimiter
                sContLine = sContLine & PoRecDetail.Item(lKeyC).Item(PoRecDetailTitles.Item(lKeyT))
            Next
            If Len(sCont) Then sCont = sCont & sLineFeedCode
            sCont = sCont & sContLine
        Next
        
        '�ҏW���ʂ�ԋp
        OutputReportInTsvFormat = sSum & sLineFeedCode & sHeader & sLineFeedCode & sCont
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : isAllOk()
    'Overview                    : �S�P�̃e�X�g�������������ǂ�����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True,Flase
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function isAllOk( _
        )
        isAllOk = (PoRecDetail.Count=func_CountCaseAs(True))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_GetDateInMilliseconds()
    'Overview                    : �������~���b�Ŏ擾����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_GetDateInMilliseconds( _
        byVal adtDate _
        , byVal adtTimer _
        )
        Dim dtNowTime        '���ݎ���
        Dim lHour            '��
        Dim lngMinute        '��
        Dim lngSecond        '�b
        Dim lngMilliSecond   '�~���b

        dtNowTime = adtTimer
        lMilliSecond = dtNowTime - Fix(dtNowTime)
        lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
        dtNowTime = Fix(dtNowTime)
        lSecond = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lMinute = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lHour = Right("0" & dtNowTime, 2)

        func_GetDateInMilliseconds = adtDate & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CountCaseAs()
    'Overview                    : ���ʂ��ƂɃP�[�X���𐔂���
    'Detailed Description        : �H����
    'Argument
    '     aboResult              : ������Ώۂ̃P�[�X���� True,Flase
    'Return Value
    '     �P�[�X��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CountCaseAs( _
        byVal aboResult _
        )
        Dim lKey : Dim lCnt : lCnt = 0
        For lKey=1 To PoRecDetail.Count
        '���ʊi�[�p�n�b�V���}�b�v����Ώۂ̃P�[�X�𐔂���
            if PoRecDetail.Item(lKey).Item(PoRecDetailTitles.Item(3)) = aboResult Then lCnt = lCnt + 1
        Next
        func_CountCaseAs = lCnt
    End Function
    
End Class
