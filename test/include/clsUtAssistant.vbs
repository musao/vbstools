'***************************************************************************************************
'FILENAME                    : clsUtAssistant.vbs
'Overview                    : �P�̃e�X�g�p�A�V�X�^���g�N���X
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
Class clsUtAssistant
'    '�N���X���ϐ�
    Private PoRecord
    Private PoRecordTitles
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '���ʊi�[�p�n�b�V���}�b�v
        Set PoRecord = CreateObject("Scripting.Dictionary")
        '���ʏڍ׃n�b�V���}�b�v�Ɋi�[������̃^�C�g����`
        Set PoRecordTitles = CreateObject("Scripting.Dictionary")
        Call PoRecordTitles.Add(1, "Seq")
        Call PoRecordTitles.Add(2, "CaseName")
        Call PoRecordTitles.Add(3, "Result")
        Call PoRecordTitles.Add(4, "Start")
        Call PoRecordTitles.Add(5, "End")
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoRecord = Nothing
    End Sub
    
    Public Property Get Record()
        Set Record = PoRecord
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
        Dim dtStart : dtStart = func_GetDateInMilliseconds()
        On Error Resume Next
        Dim boResult : boResult = GetRef(asCaseName)
        Dim dtEnd : dtEnd = func_GetDateInMilliseconds()
        If Err.Number Or Not(boResult) Then boResult = False
        
        '���ʂ��L�^
        Dim lSeq : lSeq = PoRecord.Count+1
        Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
        With PoRecordTitles
            Call oTemp.Add(.Item(1), lSeq)
            Call oTemp.Add(.Item(2), asCaseName)
            Call oTemp.Add(.Item(3), boResult)
            Call oTemp.Add(.Item(4), dtStart)
            Call oTemp.Add(.Item(5), dtEnd)
        End With
        Call PoRecord.Add(lSeq, oTemp)
        
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
        Dim lKeyT : Dim lKeyC
        
        '�w�b�_�̕ҏW
        Dim sHeader : sHeader = ""
        For lKeyT=1 To PoRecordTitles.Count
        '�w�b�_�̓^�C�g����`�̓��e�����ɏo�͂���
            If Len(sHeader) Then sHeader = sHeader & sDelimiter
            sHeader = sHeader & PoRecordTitles.Item(lKeyT)
        Next
        
        '���e�̕ҏW
        Dim sContLine
        Dim sCont : sCont = ""
        For lKeyC=1 To PoRecord.Count
        '���e�͌��ʊi�[�p�n�b�V���}�b�v�����ɏ�������
            sContLine = ""
            For lKeyT=1 To PoRecordTitles.Count
            '���ʂ��ƂɃ^�C�g�����L�[�ɒl�����o��
                If Len(sContLine) Then sContLine = sContLine & sDelimiter
                sContLine = sContLine & PoRecord.Item(lKeyC).Item(PoRecordTitles.Item(lKeyT))
            Next
            If Len(sCont) Then sCont = sCont & sLineFeedCode
            sCont = sCont & sContLine
        Next
        
        '�ҏW���ʂ�ԋp
        OutputReportInTsvFormat = sHeader & sLineFeedCode & sCont
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : isAllOk()
    'Overview                    : �SUT�������������ǂ�����Ԃ�
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
        isAllOk = False
        Dim lKey
        For lKey=1 To PoRecord.Count
        '���ʊi�[�p�n�b�V���}�b�v�����Ɋm�F����AFalse������ΏI������
            if Not(PoRecord.Item(lKey).Item(PoRecordTitles.Item(3))) Then Exit Function
        Next
        isAllOk = True
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_GetDateInMilliseconds()
    'Overview                    : ���ݓ������~���b�Ŏ擾����
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
    Private Function func_GetDateInMilliseconds()
        Dim dtNowTime        '���ݎ���
        Dim lHour            '��
        Dim lngMinute        '��
        Dim lngSecond        '�b
        Dim lngMilliSecond   '�~���b

        dtNowTime = Timer
        lMilliSecond = dtNowTime - Fix(dtNowTime)
        lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
        dtNowTime = Fix(dtNowTime)
        lSecond = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lMinute = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lHour = Right("0" & dtNowTime, 2)

        func_GetDateInMilliseconds = Date() & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
    End Function
    
End Class
