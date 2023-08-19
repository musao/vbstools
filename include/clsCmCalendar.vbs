'***************************************************************************************************
'FILENAME                    : clsCmCalendar.vbs
'Overview                    : ���t�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmCalendar()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     ���t�N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmCalendar( _
    )
    Set new_clsCmCalendar = (New clsCmCalendar).GetNow()
End Function

Class clsCmCalendar
    '�N���X���ϐ��A�萔
    Private PdtNow
    Private PdbTimer
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        PdtNow = 0
        PdbTimer = 0
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : GetNow()
    'Overview                    : ���̓��t�������擾����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function GetNow( _
        )
        Set GetNow = func_CmCalendarGetNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : SetDateTime()
    'Overview                    : �w�肵�����t������ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avNow                  : �ݒ肷����t����
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetDateTime( _
        ByVal avNow _
        )
        Set SetDateTime = func_CmCalendarSetDate(avNow)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : SetDateTimeWithFractionalSeconds()
    'Overview                    : �w�肵�����t��������у~���b��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avNow                  : �ݒ肷����t����
    '     avTimer                : �ݒ肷��~���b�iTimer()�̒l�j
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetDateTimeWithFractionalSeconds( _
        ByVal avNow _
        , ByVal avTimer _
        )
        Set SetDateTimeWithFractionalSeconds = func_CmCalendarSetDateWithFractionalSeconds(avNow, avTimer)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : DisplayFormatAs()
    'Overview                    : ���t�𐮌`����
    'Detailed Description        : func_CmCalendarSetDisplayFormatAs()�ɈϏ�����
    'Argument
    '     asFormat               : �\���`��
    'Return Value
    '     ���`�������t
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function DisplayFormatAs( _
        ByVal asFormat _
        )
        DisplayFormatAs = func_CmCalendarSetDisplayFormatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : GetSerial()
    'Overview                    : �V���A���l��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �V���A���l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function GetSerial( _
        )
       GetSerial = CDbl(Fix(PdtNow) + PdbTimer/(60*60*24))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : DifferenceInScondsFrom()
    'Overview                    : ����b���ŕԂ�
    'Detailed Description        : �H����
    'Argument
    '     aoTarget               : ��r����clsCmCalendar�^�̃C���X�^���X
    'Return Value
    '     ���̕b��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function DifferenceInScondsFrom( _
        byRef aoTarget _
        )
        Dim dbDifference : dbDifference = CDbl((Me.GetSerial()-aoTarget.GetSerial())*60*60*24)
        DifferenceInScondsFrom = Fix(dbDifference) & "." _
                                 & func_CM_FillInTheCharacters( _
                                                              Fix( (dbDifference - Fix(dbDifference))*10^6 ) _
                                                              , 6 _
                                                              , "0" _
                                                              , False _
                                                              , True _
                                                              )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : CompareTo()
    'Overview                    : ���t�̑召��r����
    'Detailed Description        : ���L��r���ʂ�Ԃ�
    '                               0  �����Ɠ��l
    '                               -1 ������菬����
    '                               1  �������傫��
    'Argument
    '     aoTarget               : ��r����clsCmCalendar�^�̃C���X�^���X
    'Return Value
    '     ��r����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function CompareTo( _
        byRef aoTarget _
        )
        Dim SerialMe : SerialMe = Me.GetSerial()
        Dim SerialTg : SerialTg = aoTarget.GetSerial()
        Dim lResult : lResult = 0
        If (SerialMe < SerialTg) Then lResult = -1
        If (SerialMe > SerialTg) Then lResult = 1
        CompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarGetNow()
    'Overview                    : ���̓��t�������擾����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarGetNow( _
        )
        PdtNow = Now()
        PdbTimer = Timer()
        Set func_CmCalendarGetNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDate()
    'Overview                    : �w�肵�����t������ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avNow                  : �ݒ肷����t����
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDate( _
        ByVal avNow _
        )
        On Error Resume Next
        PdtNow = CDate(avNow)
        PdbTimer = 0
        If Err.Number Then
            PdtNow = 0
            Err.Clear
        End If
        Set func_CmCalendarSetDate = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDateWithFractionalSeconds()
    'Overview                    : �w�肵�����t�����ƃ~���b��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avNow                  : �ݒ肷����t����
    '     avTimer                : �ݒ肷��~���b�iTimer()�̒l�j
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDateWithFractionalSeconds( _
        ByVal avNow _
        , ByVal avTimer _
        )
        On Error Resume Next
        PdtNow = CDate(avNow)
        PdbTimer = CDbl(avTimer)
        If Err.Number Then
            PdtNow = 0
            PdbTimer = 0
            Err.Clear
        End If
        Set func_CmCalendarSetDateWithFractionalSeconds = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDisplayFormatAs()
    'Overview                    : ���t�𐮌`����
    'Detailed Description        : ���L�ݒ�l�͓��t�̐��l������A���L�ȊO�̒l�͂��̂܂܎g�p����
    '                              �Ȃ��A���t��8�̏ꍇ��"DD"��"08"�A"D"��"8"��\������
    '                              ��j "YY/M/DD hh:mm:ss.000" �� 23/1/04 16:55:12.345
    '                               YY[YY]    ����N
    '                               M{M]      ��
    '                               D{D]      ��
    '                               h{h]      ��
    '                               m{m]      ��
    '                               s{s]      �b
    '                               .000      �~���b�܂�
    '                               .000000   �i�m�b�܂�
    'Argument
    '     asFormat               : �\���`��
    'Return Value
    '     ���`�������t
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDisplayFormatAs( _
        byVal asFormat _
        )
        Dim oConversionSettings : Set oConversionSettings = CreateObject("Scripting.Dictionary")
        With oConversionSettings
            '�ϊ��e�[�u����`
            .Add "YYYY", Array("UseDatePart()", "yyyy", False)
            .Add "yyyy", Array("UseDatePart()", "yyyy", False)
            .Add "YY", Array("UseDatePart()", "yyyy", True)
            .Add "yy", Array("UseDatePart()", "yyyy", True)
            .Add "MM", Array("UseDatePart()", "m", False)
            .Add "M", Array("UseDatePart()", "m", False)
            .Add "DD", Array("UseDatePart()", "d", False)
            .Add "dd", Array("UseDatePart()", "d", False)
            .Add "D", Array("UseDatePart()", "d", False)
            .Add "d", Array("UseDatePart()", "d", False)
            .Add "HH", Array("UseDatePart()", "h", False)
            .Add "hh", Array("UseDatePart()", "h", False)
            .Add "H", Array("UseDatePart()", "h", False)
            .Add "h", Array("UseDatePart()", "h", False)
            .Add "mm", Array("UseDatePart()", "n", False)
            .Add "m", Array("UseDatePart()", "n", False)
            .Add "SS", Array("UseDatePart()", "s", False)
            .Add "ss", Array("UseDatePart()", "s", False)
            .Add "S", Array("UseDatePart()", "s", False)
            .Add "s", Array("UseDatePart()", "s", False)
            .Add ".000000", Array("GetFractionalSeconds")
            .Add ".000", Array("GetFractionalSeconds")
            
            Dim lPos : lPos=1
            Dim sResult : sResult=""
            Dim lKeyLen : Dim boIsMatch : Dim sItemValue : Dim sKey : Dim vItem
            Do Until(Len(asFormat)<lPos)
                '������
                boIsMatch = False : sItemValue = ""
                
                '���ׂĂ̕ϊ��e�[�u���̏����m�F����
                For Each sKey In .Keys
                    '�L�[�̕��������擾
                    lKeyLen=Len(sKey)
                    
                    If StrComp(sKey, Mid(asFormat, lPos, lKeyLen))=0 Then
                    '�ϊ��e�[�u���ɂ��镶���ƈ�v�����ꍇ
                        vItem = .Item(sKey)
                        If vItem(0)="UseDatePart()" Then
                        'PdtNow����DatePart()�Œl�����o���ꍇ
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtNow), lKeyLen, "0", vItem(2), True)
                        Else
                        'PdbTimer����~���b�������o���ꍇ
                            sItemValue = "." & func_CM_FillInTheCharacters(Fix((PdbTimer-Fix(PdbTimer))*10^(lKeyLen-1)), lKeyLen-1, "0", False, True)
                        End If
                        boIsMatch = True : Exit For
                    End If
                Next
                
                If boIsMatch Then
                '�ϊ��e�[�u������̏ꍇ�A�}�b�`�����L�[�̕����������i�߂�
                    lPos=lPos+lKeyLen
                Else
                '�ϊ��e�[�u���Ȃ��̏ꍇ�AasFormat��1���������̂܂܎g�p��1�����i�߂�
                    sItemValue = Mid(asFormat, lPos, 1)
                    lPos=lPos+1
                End If
                sResult = sResult & sItemValue
            Loop
            
        End With
        func_CmCalendarSetDisplayFormatAs = sResult
        Set oConversionSettings = Nothing
    End Function
    
End Class