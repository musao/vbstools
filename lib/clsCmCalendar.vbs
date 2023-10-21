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
    Private PdtDate, PdbFractionalSec, PsDefaultFormat
    
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
        PdtDate = Empty
        PdbFractionalSec = Empty
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
       If Not IsEmpty(PdbFractionalSec) Then dbFractionalSec = PdbFractionalSec/(60*60*24)
       serial = Cdbl(PdtDate) + dbFractionalSec
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let serial()
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
        PdbFractionalSec = dbSec - Fix(dbSec)
        PdtDate = Cdate(adbSerial - PdbFractionalSec/60/60/24)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : �f�t�H���g�̌`���ŕ\������
    'Detailed Description        : func_CmCalendaFormatAs()�ɈϏ�����
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
        toString = func_CmCalendaFormatAs(PsDefaultFormat)
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
        If Strcomp(TypeName(Me), TypeName(aoTarget), vbBinaryCompare)<>0 Then
            Err.Raise 438, "clsCmCalendar.vbs:clsCmCalendar+compareTo()", "�I�u�W�F�N�g�ŃT�|�[�g����Ă��Ȃ��v���p�e�B�܂��̓��\�b�h�ł��B"
            Exit Function
        End If

        Dim SerialMe : SerialMe = Me.serial()
        Dim SerialTg : SerialTg = aoTarget.serial()
        Dim lResult : lResult = 0
        If (SerialMe < SerialTg) Then lResult = -1
        If (SerialMe > SerialTg) Then lResult = 1
        compareTo = lResult
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
        If Strcomp(TypeName(Me), TypeName(aoTarget), vbBinaryCompare)<>0 Then
            Err.Raise 438, "clsCmCalendar.vbs:clsCmCalendar+differenceFrom()", "�I�u�W�F�N�g�ŃT�|�[�g����Ă��Ȃ��v���p�e�B�܂��̓��\�b�h�ł��B"
            Exit Function
        End If

        differenceFrom = math_roundDown(Me.serial()*60*60*24-aoTarget.serial()*60*60*24, 5)

    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : formatAs()
    'Overview                    : ���t�𐮌`����
    'Detailed Description        : func_CmCalendaFormatAs()�ɈϏ�����
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
        formatAs = func_CmCalendaFormatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getNow()
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
        Set getNow = func_CmCalendarGetNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setDateTime()
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
        Set setDateTime = func_CmCalendarSetDate(avDateTime)
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
        PdtDate = Now()
        
        Dim dbTimer : dbTimer = Timer()
        PdbFractionalSec = dbTimer - Fix(dbTimer)

        Set func_CmCalendarGetNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDate()
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
    Private Function func_CmCalendarSetDate( _
        ByVal avDateTime _
        )
        Dim sPtn : sPtn = "^([^.]+)\.(\d+)$"
        If new_Re(sPtn, "").Test(avDateTime) Then
            PdtDate = Cdate(new_Re(sPtn, "").Replace(avDateTime, "$1"))
            PdbFractionalSec = Cdbl("0." & new_Re(sPtn, "").Replace(avDateTime, "$2"))
        Else
            PdtDate = Cdate(avDateTime)
            PdbFractionalSec = Empty
        End If
        Set func_CmCalendarSetDate = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendaFormatAs()
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
    Private Function func_CmCalendaFormatAs( _
        byVal asFormat _
        )
        Dim oConversionSettings : Set oConversionSettings = new_Dic()
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
            .Add "000000", Array("GetFractionalSeconds")
            .Add "000", Array("GetFractionalSeconds")
            
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
                        'PdtDate����DatePart()�Œl�����o���ꍇ
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDate), lKeyLen, "0", vItem(2), True)
                        Else
                        '�b���̏����������o���ꍇ
                            Dim dbFractionalSec : dbFractionalSec =0
                            If Not IsEmpty(PdbFractionalSec) Then dbFractionalSec = PdbFractionalSec
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
        func_CmCalendaFormatAs = sResult
        Set oConversionSettings = Nothing
    End Function
    
End Class
