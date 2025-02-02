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
Class clsCmCalendar
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
       Dim dbFractionalSec : dbFractionalSec = 0
       If Not IsNull(PdbElapsedSeconds) Then dbFractionalSec = PdbElapsedSeconds/(60*60*24)
       serial = Cdbl(PdtDateTime) + dbFractionalSec
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let serial() ->���p�~�\��
    'Overview                    : ���t�̃V���A���l��ݒ�
    'Detailed Description        : �V���A���l�Ƃ�1900/1/1��1�Ƃ��āA�����o�߂��������������l
    'Argument
    '     adbSerial              : ���t�̃V���A���l
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let serial( _
        byVal adbSerial _
        )
        Dim dbSec : dbSec = (adbSerial - Fix(adbSerial))*60*60*24
        PdbElapsedSeconds = dbSec - Fix(dbSec)
        PdtDateTime = Cdate(adbSerial - PdbElapsedSeconds/60/60/24)
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
        ast_argsIsSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+compareTo()", "That object is not a calendar class."
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
        ast_argsIsSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+differenceFrom()", "That object is not a calendar class."
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
        ByVal asFormat _
        )
        formatAs = this_formatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getNow() ->���p�~�\��
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
    Public Function getNow( _
        )
        Set getNow = this_getNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setDateTime() ->���p�~�\��
    'Overview                    : �w�肵�����t������ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avDateTime             : �ݒ肷����t����
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setDateTime( _
        ByVal avDateTime _
        )
        Set setDateTime = this_setDate(avDateTime)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : newInstance()
    'Overview                    : �C���X�^���X���쐬����
    'Detailed Description        : this_newInstance()�ɈϏ�����
    'Argument
    '     adtDateTime            : ����
    '     adbElapsedSeconds      : �o�ߕb
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function newInstance( _
        ByVal adtDateTime _
        , ByVal adbElapsedSeconds _
        )
        Set newInstance = this_newInstance(adtDateTime, adbElapsedSeconds, TypeName(Me)&"+newInstance()")
    End Function
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getNow() ->���p�~�\��
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
    Private Function this_getNow( _
        )
        PdtDateTime = Now()
        
        Dim dbTimer : dbTimer = Timer()
        PdbElapsedSeconds = dbTimer - Fix(dbTimer)

        Set this_getNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setDate() ->���p�~�\��
    'Overview                    : �w�肵�����t������ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     avDateTime             : �ݒ肷����t����
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setDate( _
        ByVal avDateTime _
        )
        Dim sPtn : sPtn = "^([^.]+)\.(\d+)$"
        If new_Re(sPtn, "").Test(avDateTime) Then
            PdtDateTime = Cdate(new_Re(sPtn, "").Replace(avDateTime, "$1"))
            PdbElapsedSeconds = Cdbl("0." & new_Re(sPtn, "").Replace(avDateTime, "$2"))
        Else
            PdtDateTime = Cdate(avDateTime)
            PdbElapsedSeconds = Null
        End If
        Set this_setDate = Me
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
                    
                    If StrComp(sKey, Mid(asFormat, lPos, lKeyLen))=0 Then
                    '�ϊ��e�[�u���ɂ��镶���ƈ�v�����ꍇ
                        vItem = .Item(sKey)
                        If cf_isSame(Cl_USE_DATAPART, vItem(0)) Then
                        'PdtDate����DatePart()�Œl�����o���ꍇ
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDateTime), lKeyLen, "0", vItem(2), True)
                        Else
                        '�b���̏����������o���ꍇ
                            Dim dbFractionalSec : dbFractionalSec =0
                            If Not IsNull(PdbElapsedSeconds) Then dbFractionalSec = PdbElapsedSeconds
                            sItemValue = func_CM_FillInTheCharacters(Fix(dbFractionalSec*10^lKeyLen), lKeyLen, "0", False, True)
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
    'Function/Sub Name           : this_newInstance()
    'Overview                    : �C���X�^���X���쐬����
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
    Private Function this_newInstance( _
        byVal adtDateTime _
        , byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argFalse IsNull(PdtDateTime), asSource, "Because it is an immutable variable, its value cannot be changed."
        this_setDateTime adtDateTime, asSource
        this_setElapsedSeconds adbElapsedSeconds, asSource
        Set this_newInstance = Me
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
        ast_argTrue IsDate(PdtDateTime), asSource, "DateTime is not a date/time."
        PadtDateTime = adtDateTime
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
        PadbElapsedSeconds = adbElapsedSeconds
    End Sub
    
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
        Dim lResult : lResult = 0

        If (Me.dateTime < aoTarget.dateTime) Then lResult = -1
        If (Me.dateTime > aoTarget.dateTime) Then lResult = 1
        If lResult <> 0 Then
            this_compareTo = lResult
            Exit Function
        End If
        
        If (Me.elapsedSeconds < aoTarget.elapsedSeconds) Then lResult = -1
        If (Me.elapsedSeconds > aoTarget.elapsedSeconds) Then lResult = 1

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

        Dim dbDiffElapsedSeconds : dbDiffElapsedSeconds = Me.elapsedSeconds - aoTarget.elapsedSeconds
        If (Me.dateTime <> aoTarget.dateTime) Then dbDiffElapsedSeconds = dbDiffElapsedSeconds+(Me.dateTime-aoTarget.dateTime)*60*60*24
        this_differenceFrom = math_roundDown(dbDiffElapsedSeconds, 5)

'        this_differenceFrom = math_roundDown(Me.serial()*60*60*24-aoTarget.serial()*60*60*24, 5)
    End Function
    
End Class
