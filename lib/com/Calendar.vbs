'***************************************************************************************************
'FILENAME                    : Calendar.vbs
'Overview                    : ���t�N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Class Calendar
    '�N���X���ϐ��A�萔
    Private PdtDateTime, PdbElapsedSeconds, PsDefaultFormat
    
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
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PdtDateTime = Null
        PdbElapsedSeconds = Null
        PsDefaultFormat = "YYYY/MM/DD hh:mm:ss.000"
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
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get dateTime()
    'Overview                    : ������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get dateTime()
       dateTime = PdtDateTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get fractionalPartOfElapsedSeconds()
    'Overview                    : �o�ߕb�̏�������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�ߕb�̏�����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get fractionalPartOfElapsedSeconds()
       fractionalPartOfElapsedSeconds = Null
       If Not IsNull(PdtDateTime) Then fractionalPartOfElapsedSeconds = this_getfractionalPartOfElapsedSeconds()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get elapsedSeconds()
    'Overview                    : �o�ߕb��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�ߕb
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get elapsedSeconds()
       elapsedSeconds = PdbElapsedSeconds
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get serial()
    'Overview                    : ���t�̃V���A���l��Ԃ�
    'Detailed Description        : �V���A���l�Ƃ�1900/1/1��1�Ƃ��āA�����o�߂��������������l
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���t�̃V���A���l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get serial()
       serial = Null
       If Not IsNull(PdtDateTime) Then serial = Cdbl(PdtDateTime)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : �f�t�H���g�̌`���ŕ\������
    'Detailed Description        : this_formatAs()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �f�t�H���g�̌`���ɐ��`�������t
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get toString()
        toString = this_formatAs(PsDefaultFormat)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : clone()
    'Overview                    : ���g�Ɠ������e�̐V�����C���X�^���X�����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �V�����C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function clone( _
        )
        Set clone = this_clone()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
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
    Public Function compareTo( _
        byRef aoTarget _
        )
        ast_argsAreSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+compareTo()", "That object is not a calendar class."
        compareTo = this_compareTo(aoTarget)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : differenceFrom()
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
    Public Function differenceFrom( _
        byRef aoTarget _
        )
        ast_argsAreSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+differenceFrom()", "That object is not a calendar class."
        differenceFrom = this_differenceFrom(aoTarget)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : formatAs()
    'Overview                    : ���t�𐮌`����
    'Detailed Description        : this_formatAs()�ɈϏ�����
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
    Public Function formatAs( _
        byVal asFormat _
        )
        formatAs = this_formatAs(asFormat)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : of()
    'Overview                    : �����ɉ������C���X�^���X���쐬����
    'Detailed Description        : this_of()�ɈϏ�����
    'Argument
    '     avArgument             : ����
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function of( _
        byRef avArgument _
        )
        Set of = this_of(avArgument, TypeName(Me)&"+of()")
    End Function
     
    '***************************************************************************************************
    'Function/Sub Name           : ofNow()
    'Overview                    : ���̓��t�������擾����
    'Detailed Description        : this_setData()�ɈϏ�����
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
    Public Function ofNow( _
        )
        Set ofNow = this_setData(Now(), Timer(), TypeName(Me)&"+ofNow()")
    End Function
       



    
    '***************************************************************************************************
    'Function/Sub Name           : this_clone()
    'Overview                    : ���g�Ɠ������e�̐V�����C���X�^���X�����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �V�����C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_clone( _
        )
        Dim oNewIns : Set oNewIns = new Calendar
        If IsNull(PdtDateTime) Then
        Else
            If IsNull(PdbElapsedSeconds) Then
                Call oNewIns.of(Array(PdtDateTime))
            Else
                Call oNewIns.of(Array(PdtDateTime, PdbElapsedSeconds))
            End If
        End If
        Set this_clone = oNewIns
        Set oNewIns = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_compareTo()
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
    '2025/02/01         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_compareTo( _
        byRef aoTarget _
        )
        this_compareTo = 0
        If IsNull(PdtDateTime) And IsNull(aoTarget.dateTime) Then Exit Function
        
        Dim lResult : lResult = 0
        If IsNull(PdtDateTime) Or (PdtDateTime < aoTarget.dateTime) Then lResult = -1
        If IsNull(aoTarget.dateTime) Or (PdtDateTime > aoTarget.dateTime) Then lResult = 1
        If lResult <> 0 Then
            this_compareTo = lResult
            Exit Function
        End If
        
        If (this_getfractionalPartOfElapsedSeconds < aoTarget.fractionalPartOfElapsedSeconds) Then lResult = -1
        If (this_getfractionalPartOfElapsedSeconds > aoTarget.fractionalPartOfElapsedSeconds) Then lResult = 1
        this_compareTo = lResult

    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_differenceFrom()
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
    '2025/02/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function this_differenceFrom( _
        byRef aoTarget _
        )
        If this_compareTo(aoTarget)=0 Then
            this_differenceFrom = 0
            Exit function
        End If

        Dim dbResult : dbResult = 0
        If IsNull(PdtDateTime) Then dbResult = -1 * ((aoTarget.dateTime)*60*60*24 + aoTarget.fractionalPartOfElapsedSeconds)
        If IsNull(aoTarget.dateTime) Then dbResult = PdtDateTime*60*60*24 + this_getfractionalPartOfElapsedSeconds
        If dbResult <> 0 Then
            this_differenceFrom = dbResult
            Exit Function
        End If

        Dim dbDiffElapsedSeconds
        dbDiffElapsedSeconds = this_getfractionalPartOfElapsedSeconds-aoTarget.fractionalPartOfElapsedSeconds

        If (PdtDateTime <> aoTarget.dateTime) Then dbDiffElapsedSeconds = dbDiffElapsedSeconds+(PdtDateTime-aoTarget.dateTime)*60*60*24
        this_differenceFrom = math_round(dbDiffElapsedSeconds, 6)

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_formatAs()
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
    '                               000       �~���b�܂�
    '                               000000    �}�C�N���b�܂�
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
    Private Function this_formatAs( _
        byVal asFormat _
        )
        this_formatAs = "<"&TypeName(Me)&">"&cf_toString(Null)
        If IsNull(PdtDateTime) Then Exit Function

        Const Cl_USE_DATAPART = 0
        Const Cl_USE_FRACTIONAL_SECONDS = 1
        With new_Dic()
            '�ϊ��e�[�u����`
            .Add "YYYY", Array(Cl_USE_DATAPART, "yyyy", False)
            .Add "yyyy", Array(Cl_USE_DATAPART, "yyyy", False)
            .Add "YY", Array(Cl_USE_DATAPART, "yyyy", True)
            .Add "yy", Array(Cl_USE_DATAPART, "yyyy", True)
            .Add "MM", Array(Cl_USE_DATAPART, "m", False)
            .Add "M", Array(Cl_USE_DATAPART, "m", False)
            .Add "DD", Array(Cl_USE_DATAPART, "d", False)
            .Add "dd", Array(Cl_USE_DATAPART, "d", False)
            .Add "D", Array(Cl_USE_DATAPART, "d", False)
            .Add "d", Array(Cl_USE_DATAPART, "d", False)
            .Add "HH", Array(Cl_USE_DATAPART, "h", False)
            .Add "hh", Array(Cl_USE_DATAPART, "h", False)
            .Add "H", Array(Cl_USE_DATAPART, "h", False)
            .Add "h", Array(Cl_USE_DATAPART, "h", False)
            .Add "mm", Array(Cl_USE_DATAPART, "n", False)
            .Add "m", Array(Cl_USE_DATAPART, "n", False)
            .Add "SS", Array(Cl_USE_DATAPART, "s", False)
            .Add "ss", Array(Cl_USE_DATAPART, "s", False)
            .Add "S", Array(Cl_USE_DATAPART, "s", False)
            .Add "s", Array(Cl_USE_DATAPART, "s", False)
            .Add "000000", Array(Cl_USE_FRACTIONAL_SECONDS)
            .Add "000", Array(Cl_USE_FRACTIONAL_SECONDS)
            
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
                    
                    If cf_isSame(sKey, Mid(asFormat, lPos, lKeyLen)) Then
                    '�ϊ��e�[�u���ɂ��镶���ƈ�v�����ꍇ
                        vItem = .Item(sKey)
                        If cf_isSame(Cl_USE_DATAPART, vItem(0)) Then
                        'PdtDate����DatePart()�Œl�����o���ꍇ
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDateTime), lKeyLen, "0", vItem(2), True)
                        Else
                        '�b���̏����������o���ꍇ
                            sItemValue = func_CM_FillInTheCharacters(math_tranc(this_getfractionalPartOfElapsedSeconds*10^lKeyLen), lKeyLen, "0", False, True)
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
        this_formatAs = sResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getfractionalPartOfElapsedSeconds()
    'Overview                    : �o�ߕb�̏�������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�ߕb�̏�����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getfractionalPartOfElapsedSeconds( _
        )
        Dim dbFract : dbFract = 0
        If Not IsNull(PdbElapsedSeconds) Then dbFract = math_round(math_fractional(PdbElapsedSeconds),7)
        this_getfractionalPartOfElapsedSeconds = dbFract
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_of()
    'Overview                    : �����ɉ������C���X�^���X���쐬����
    'Detailed Description        : this_setData()�ɈϏ�����
    '                              �ȉ��̓��͌������s��
    '                              1.�z��łȂ��ꍇ
    '                                Date�^�i�����_�ȉ��̕b���������Ă��悢�j
    '                              2.�z��̏ꍇ�͗v�f���ɉ������`�F�b�N���s��
    '                                1-1.�v�f����1��
    '                                 e(0) -> Date�^�i�����_�ȉ��̕b���������Ă��悢�j
    '                                1-2.�v�f����2��
    '                                 e(0) -> Date�^
    '                                 e(1) -> Double�^
    '                                1-3.�v�f����6��
    '                                 e(0-5) -> "e(0)/e(1)/e(2) e(3):e(4):e(5)"��Date�^
    '                                1-4.�v�f����7��
    '                                 e(0-5) -> "e(0)/e(1)/e(2) e(3):e(4):e(5)"��Date�^
    '                                 e(6) -> Double�^
    '                                1-5.��L�ȊO�̗v�f���̓G���[�Ƃ���
    'Argument
    '     avArgument             : ����
    '     asSource               : �\�[�X
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_of( _
        byRef avArgument _
        , byVal asSource _
        )
        Dim dtDateTime, dbElapsedSeconds, boIsError
        dtDateTime = Null
        dbElapsedSeconds = Null
        boIsError = False
        
        On Error Resume Next
        If Not(IsArray(avArgument)) Then
        '�z��łȂ��ꍇ
            Call this_ofForOneArg(avArgument, dtDateTime, dbElapsedSeconds)
        ElseIf new_Arr().hasElement(avArgument) Then
        '�z��̗v�f������ꍇ
            Dim e : e = avArgument
            Select Case Ubound(e)
                Case 0:
                    Call this_ofForOneArg(e(0), dtDateTime, dbElapsedSeconds)
                Case 1:
                    dtDateTime = Cdate(e(0))
                    dbElapsedSeconds = Cdbl(e(1))
                Case 5:
                    dtDateTime = Cdate(e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5))
                Case 6:
                    dtDateTime = Cdate(e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5))
                    dbElapsedSeconds = Cdbl(e(6))
            End Select
        End If
        If Err.Number<>0 Then boIsError=True
        On Error Goto 0

        ast_argFalse boIsError, asSource, "invalid argument. " & cf_toString(avArgument)

        Set this_of = this_setData(dtDateTime, dbElapsedSeconds, asSource)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_ofForOneArg()
    'Overview                    : ���t�^�ɕϊ�����
    'Detailed Description        : �H����
    'Argument
    '     avDateTime             : �����̓��t����
    '     adtDateTime            : ����
    '     dbElapsedSeconds       : �o�ߕb
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_ofForOneArg( _
        byRef avDateTime _
        , byRef adtDateTime _
        , byRef adbElapsedSeconds _
        )
        Dim oRe : Set oRe = new_Re("^([^.]+)\.(\d+)$", "")
        If oRe.Test(avDateTime) Then
            adtDateTime = Cdate(oRe.Replace(avDateTime, "$1"))
            Dim dbElapsedSecondsByDt : dbElapsedSecondsByDt = math_tranc(math_fractional(adtDateTime)*24*60*60)
            adbElapsedSeconds = dbElapsedSecondsByDt + Cdbl("0." & oRe.Replace(avDateTime, "$2"))
        Else
            adtDateTime = Cdate(avDateTime)
        End If
        Set oRe = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : �f�[�^��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     adtDateTime            : ����
    '     adbElapsedSeconds      : �o�ߕb
    '     asSource               : �\�[�X
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setData( _
        byVal adtDateTime _
        , byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argNull PdtDateTime, asSource, "Because it is an immutable variable, its value cannot be changed."
        this_setDateTime adtDateTime, asSource
        this_setElapsedSeconds adbElapsedSeconds, asSource
        Set this_setData = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_setDateTime()
    'Overview                    : PadtDateTime�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     adtDateTime            : ����
    '     asSource               : �\�[�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setDateTime( _
        byVal adtDateTime _
        , byVal asSource _
        )
        ast_argTrue IsDate(adtDateTime), asSource, "DateTime is not a date/time."
        PdtDateTime = Cdate(adtDateTime)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setElapsedSeconds()
    'Overview                    : PadbElapsedSeconds�̃Z�b�^�[
    'Detailed Description        : �H����
    'Argument
    '     adbElapsedSeconds      : �o�ߕb
    '     asSource               : �\�[�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setElapsedSeconds( _
        byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argTrue (IsNull(adbElapsedSeconds) Or cf_isNonNegativeNumber(adbElapsedSeconds)), asSource, "ElapsedSeconds must be null or a non-negative number."
        If Not(IsNull(adbElapsedSeconds)) Then PdbElapsedSeconds = Cdbl(adbElapsedSeconds)
    End Sub
    
End Class
