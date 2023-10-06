'***************************************************************************************************
'FILENAME                    : clsCmBufferedReader.vbs
'Overview                    : �t�@�C���Ǎ��o�b�t�@�����O�����N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmBufferedReader()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     aoTextStream           : �e�L�X�g�X�g���[���I�u�W�F�N�g
'Return Value
'     �t�@�C���Ǎ��o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmBufferedReader( _
    byRef aoTextStream _
    )
    Set new_clsCmBufferedReader = (New clsCmBufferedReader).SetTextStream(aoTextStream)
End Function

Class clsCmBufferedReader
    '�N���X���ϐ��A�萔
    Private PoTextStream, PoOutbound, PoInbound, PoBuffer, PlReadSize
    
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
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PlReadSize = 5000                 '�f�t�H���g��5000�o�C�g
        Set PoTextStream = Nothing
        Dim vArr : vArr = Array("Line", Empty, "Column", Empty, "AtEndOfLine", Empty, "AtEndOfStream", Empty)
        Set PoOutbound = new_DictSetValues(vArr)
        Set PoInbound = new_DictSetValues(vArr)
        Set PoBuffer = new_DictSetValues(Array("Buffer", Empty, "Pointer", Empty, "Length", Empty))
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : �f�X�g���N�^
    'Detailed Description        : �o�b�t�@�̖��o�͕����o�͂��Ă���I���������s��
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Call sub_CmBufferedReaderClose()
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let ReadSize()
    'Overview                    : �Ǎ��T�C�Y��ݒ肷��
    'Detailed Description        : �Ǎ��v�����ɓǍ��o�b�t�@�̃T�C�Y������𒴂����ꍇ
    '                              �t�@�C����Ǎ���
    'Argument
    '     alReadSize             : �Ǎ��T�C�Y
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let ReadSize( _
        byVal alReadSize _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alReadSize, 2) Then
            PlReadSize = CLng(alReadSize)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ReadSize()
    'Overview                    : �Ǎ��T�C�Y��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ǎ��T�C�Y
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ReadSize()
        ReadSize = PlReadSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get TextStream()
    'Overview                    : �e�L�X�g�X�g���[����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �e�L�X�g�X�g���[���I�u�W�F�N�g
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get TextStream()
        Set TextStream = aoTextStream
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Line()
    'Overview                    : ���݂̍s�ԍ���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���݂̍s�ԍ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Line()
        Line = PoOutbound.Item("Line")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Column()
    'Overview                    : ���݂̗�ԍ���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���݂̗�ԍ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Column()
        Column = PoOutbound.Item("Column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get AtEndOfStream()
    'Overview                    : �s���̏ꍇ��True��Ԃ�
    'Detailed Description        : ���s���}�[�J�[�̒��O�Ƀt�@�C�� �|�C���^�[������ꍇ�� true ��Ԃ��A
    '                              �����łȂ��ꍇ�� false ��Ԃ��܂��B
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True:�s�� / False:�s���ȊO
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get AtEndOfStream()
        AtEndOfStream = PoOutbound.Item("AtEndOfStream")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get AtEndOfLine()
    'Overview                    : �t�@�C���̏I�[�̏ꍇ��True��Ԃ�
    'Detailed Description        : �Ō�Ƀt�@�C�� �|�C���^�[������ꍇ�� true ��Ԃ��A
    '                              �����łȂ��ꍇ�� false ��Ԃ��܂��B
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True:�t�@�C���̏I�[ / False:�t�@�C���̏I�[�ȊO
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get AtEndOfLine()
        AtEndOfLine = PoOutbound.Item("AtEndOfLine")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : SetTextStream()
    'Overview                    : �e�L�X�g�X�g���[����ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoTextStream           : �e�L�X�g�X�g���[���I�u�W�F�N�g
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetTextStream( _
        byRef aoTextStream _
        )
        Set PoTextStream = aoTextStream
        Set SetTextStream = Me
        
        'Inbound�AOutbound�Ȃǂ̏�������������
        Call sub_CmBufferedReaderInitialize()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Read()
    'Overview                    : �t�@�C������w�肵�������������ǂݍ���
    'Detailed Description        : func_CmBufferedReaderRead()�ɈϏ�����
    'Argument
    '     alLength               : �ǂݍ��ޕ�����
    'Return Value
    '     �ǂݍ��񂾕�����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Read( _
        byVal alLength _
        )
        Read = func_CmBufferedReaderRead(alLength)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReadLine()
    'Overview                    : �t�@�C������1�s�ǂݍ���
    'Detailed Description        : func_CmBufferedReaderReadLine()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ǂݍ��񂾕�����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function ReadLine( _
        )
        ReadLine = func_CmBufferedReaderReadLine()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReadAll()
    'Overview                    : �t�@�C���S�̂�ǂݍ���
    'Detailed Description        : func_CmBufferedReaderReadAll()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ǂݍ��񂾕�����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function ReadAll( _
        )
        ReadAll = func_CmBufferedReaderReadAll()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Skip()
    'Overview                    : �t�@�C������w�肵�������������X�L�b�v����
    'Detailed Description        : func_CmBufferedReaderRead()�ɈϏ�����
    'Argument
    '     alLength               : �X�L�b�v���镶����
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Skip( _
        byVal alLength _
        )
        Call func_CmBufferedReaderRead(alLength)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : SkipLine()
    'Overview                    : �t�@�C������1�s�X�L�b�v����
    'Detailed Description        : func_CmBufferedReaderReadLine()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub SkipLine( _
        )
        Call func_CmBufferedReaderReadLine()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Close()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : sub_CmBufferedReaderClose()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Close( _
        )
        Call sub_CmBufferedReaderClose()
    End Sub
    
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderRead()
    'Overview                    : �t�@�C������w�肵�������������ǂݍ���
    'Detailed Description        : �H����
    'Argument
    '     alLength               : �ǂݍ��ޕ�����
    'Return Value
    '     �ǂݎ����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmBufferedReaderRead( _
        byVal alLength _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�o�b�t�@���Ȃ���΃t�@�C����ǂݎ��
        Do While PoInbound.Item("AtEndOfStream")=False And (PoBuffer.Item("Length")-PoBuffer.Item("Pointer")+1)<alLength
        '�C���o�E���h���ǂݏo���\�iAtEndOfStream=False�j���o�b�t�@�̖��ǂݏo�������̒������ǂݍ��ޕ����������̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            Call func_CmBufferedReaderReadFile(False)
        Loop
        
        '�o�b�t�@����w�肵�����������o��
        Dim sRet : sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), alLength)
        
        '�|�C���^���X�V
        PoBuffer.Item("Pointer") = PoBuffer.Item("Pointer")+Len(sRet)
        
        '�A�E�g�o�E���h�̏����X�V
        Dim oArr : Set oArr = new_ArraySplit(Mid(PoBuffer.Item("Buffer"), 1, PoBuffer.Item("Pointer") - 1), vbLf)
        oArr.Reverse()
        With PoOutbound
            .Item("Line") = oArr.Length
            .Item("Column") = Len(oArr(0))+1
            .Item("AtEndOfStream") = PoInbound.Item("AtEndOfStream") And (PoBuffer.Item("Pointer") > PoBuffer.Item("Length"))
            .Item("AtEndOfLine") = .Item("AtEndOfStream") Or (StrComp(Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), 1), vbLf, vbBinaryCompare)=0)
        End With
        
        '�߂�l��Ԃ�
        func_CmBufferedReaderRead = sRet
        Set oArr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadLine()
    'Overview                    : �t�@�C������1�s�ǂݍ���
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ǂݎ����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmBufferedReaderReadLine( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�o�b�t�@���Ȃ���΃t�@�C����ǂݎ��
        Do While PoInbound.Item("AtEndOfStream")=False And InStr(PoBuffer.Item("Pointer"), PoBuffer.Item("Buffer"), vbLf, vbBinaryCompare)=0
        '�C���o�E���h���ǂݏo���\�iAtEndOfStream=False�j���|�C���^�̂���s���o�b�t�@�̍ŏI�s�̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            Call func_CmBufferedReaderReadFile(False)
        Loop
        
        '�s���ivbLf�j����������
        Dim lPosRowEnd : lPosRowEnd = InStr(PoBuffer.Item("Pointer"), PoBuffer.Item("Buffer"), vbLf, vbBinaryCompare)
        Dim sRet
        If lPosRowEnd=0 Then
        '�s���ivbLf�j��������Ȃ��������t�@�C���̏I�[�̏ꍇ
            '�|�C���^�ȍ~�S�Ă̕�����Ԃ�
            sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"))
            '�t�@�C���̏I�[�Ƀ|�C���^���X�V
            PoBuffer.Item("Pointer") = PoBuffer.Item("Length")+1
        Else
        '�s���ivbLf�j�������������t�@�C���̏I�[�łȂ��ꍇ
            '�|�C���^���玟�̉��s�����ivbLf�j�܂Łi���s�������܂܂Ȃ��j��Ԃ�
            sRet = Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), lPosRowEnd-PoBuffer.Item("Pointer"))
           '�ŌオvbCr�̏ꍇ�͍폜����
           If StrComp(Right(sRet, 1), vbCr, vbBinaryCompare)=0 Then sRet = Mid(sRet, 1, Len(sRet)-1)
           '���̍s�̍s���i���������s���̈ʒu+1�j�Ƀ|�C���^���X�V
           PoBuffer.Item("Pointer") = lPosRowEnd+1
        End If
        
        '�A�E�g�o�E���h�̏����X�V
        With PoOutbound
            Dim boEof : If .Item("Line")+1>PoInbound.Item("Line") Then boEof = True Else boEof = False
            If boEof Then
            '�t�@�C���̏I�[�܂œǂݏo�����ꍇ
                '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
                Call sub_CmBufferedReaderCopyInboundStateToOutbound()
            Else
            '�t�@�C���̏I�[�܂œǂݏo���ĂȂ��ꍇ
                '���̍s�̍s���ɍX�V����
                .Item("Line") = .Item("Line")+1
                .Item("Column") = 1
                .Item("AtEndOfStream") = False
                .Item("AtEndOfLine") = (StrComp(Mid(PoBuffer.Item("Buffer"), PoBuffer.Item("Pointer"), 1), vbLf, vbBinaryCompare)=0)
            End If
        End With
        
        '�߂�l��Ԃ�
        func_CmBufferedReaderReadLine = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadAll()
    'Overview                    : �t�@�C���S�̂�ǂݎ��
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ǂݎ����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmBufferedReaderReadAll( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�t�@�C���S�̂�ǂݎ��
        Dim sRet : sRet = func_CmBufferedReaderReadFile(True)
        
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        Call sub_CmBufferedReaderCopyInboundStateToOutbound()
        '�|�C���^���X�V
        PoBuffer.Item("Pointer") = PoBuffer.Item("Length")+1
        
        '�߂�l��Ԃ�
        func_CmBufferedReaderReadAll = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedReaderReadFile()
    'Overview                    : �w�肵�����@�Ńt�@�C����ǂݍ���Ńo�b�t�@�ɏ�������
    'Detailed Description        : �Ǎ��񂾌�ɃC���o�E���h�̏�Ԃ��擾����
    'Argument
    '     aboIsReadAll           : �t�@�C���̓ǂݎ����@
    '                                True :�t�@�C���S�̂�ǂݎ��
    '                                False:�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
    'Return Value
    '     �ǂݎ����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmBufferedReaderReadFile( _
        byVal aboIsReadAll _
        )
        
        '�t�@�C����Ǎ���
        Dim sText : sText = ""
        If aboIsReadAll Then
            sText = PoTextStream.ReadAll
        Else
            sText = PoTextStream.Read(PlReadSize)
        End If
        '�o�b�t�@�̍X�V
        With PoBuffer
            .Item("Buffer") = .Item("Buffer") & sText
            .Item("Length") = Len(.Item("Buffer"))
        End With
        '�C���o�E���h�̏�Ԃ��擾����
        Call sub_CmBufferedReaderGetInboundStatus()
        '�߂�l��Ԃ�
        func_CmBufferedReaderReadFile = sText
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderClose()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedReaderClose( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        Call PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderGetInboundStatus()
    'Overview                    : �C���o�E���h�̏�Ԃ��擾����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedReaderGetInboundStatus( _
        )
        With PoTextStream
            '�C���o�E���h�̏�Ԃ��擾����
            Set PoInbound = new_DictSetValues(Array("Line", .Line, "Column", .Column, "AtEndOfLine", .AtEndOfLine, "AtEndOfStream", .AtEndOfStream))
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderCopyInboundStateToOutbound()
    'Overview                    : �C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedReaderCopyInboundStateToOutbound( _
        )
        With PoInbound
            '�A�E�g�o�E���h�̏�ԂɃC���o�E���h�̏�Ԃ��R�s�[����
            Dim sKey, oOutbound
            Set oOutbound = new_Dictionary()
            For Each sKey In Array("Line", "Column", "AtEndOfLine", "AtEndOfStream")
                oOutbound.Add sKey, .Item(sKey)
            Next
        End With
        Set PoOutbound = oOutbound
        Set oOutbound = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedReaderInitialize()
    'Overview                    : Inbound�AOutbound�Ȃǂ̏�������������
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedReaderInitialize( _
        )
        '�C���o�E���h�̏�Ԃ��擾����
        Call sub_CmBufferedReaderGetInboundStatus()
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        Call sub_CmBufferedReaderCopyInboundStateToOutbound()
        '�|�C���^�̏�����
        Set PoBuffer = new_DictSetValues(Array("Pointer", 1, "Buffer", "", "Length", 0))
    End Sub
    
End Class
