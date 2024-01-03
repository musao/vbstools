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
        Dim vArr : vArr = Array("line", Empty, "column", Empty, "atEndOfLine", Empty, "atEndOfStream", Empty)
        Set PoOutbound = new_DicWith(vArr)
        Set PoInbound = new_DicWith(vArr)
        Set PoBuffer = new_DicWith(Array("buffer", Empty, "pointer", Empty, "length", Empty))
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
        sub_CmBufferedReaderClose
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let readSize()
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
    Public Property Let readSize( _
        byVal alReadSize _
        )
        Dim boFlg : boFlg = False
        If cf_isNumeric(alReadSize) Then
            If CDbl(alReadSize)>0 Then
                boFlg = True
            End If
        End If
        
        If boFlg Then
            PlReadSize = CDbl(alReadSize)
        Else
            Err.Raise 1031, "clsCmBufferedReader.vbs:clsCmBufferedReader+readSize()", "�s���Ȑ����ł��B"
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get readSize()
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
    Public Property Get readSize()
        readSize = PlReadSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get textStream()
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
    Public Property Get textStream()
        Set textStream = PoTextStream
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get line()
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
    Public Property Get line()
        line = PoOutbound.Item("line")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get column()
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
    Public Property Get column()
        column = PoOutbound.Item("column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get atEndOfStream()
    'Overview                    : �s���̏ꍇ��True��Ԃ�
    'Detailed Description        : ���s���}�[�J�[�̒��O�Ƀt�@�C�� �|�C���^�[������ꍇ�� true ��Ԃ��A
    '                              �����łȂ��ꍇ�� false ��Ԃ��B
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
    Public Property Get atEndOfStream()
        atEndOfStream = PoOutbound.Item("atEndOfStream")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get atEndOfLine()
    'Overview                    : �t�@�C���̏I�[�̏ꍇ��True��Ԃ�
    'Detailed Description        : �Ō�Ƀt�@�C�� �|�C���^�[������ꍇ�� true ��Ԃ��A
    '                              �����łȂ��ꍇ�� false ��Ԃ��B
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
    Public Property Get atEndOfLine()
        atEndOfLine = PoOutbound.Item("atEndOfLine")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : setTextStream()
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
    Public Function setTextStream( _
        byRef aoTextStream _
        )
        If Not func_CM_UtilIsTextStream(aoTextStream) Then
            Err.Raise 438, "clsCmBufferedReader.vbs:clsCmBufferedReader+setTextStream()", "�I�u�W�F�N�g�ŃT�|�[�g����Ă��Ȃ��v���p�e�B�܂��̓��\�b�h�ł��B"
            Exit Function
        End If

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
        'Inbound�AOutbound�Ȃǂ̏�������������
        sub_CmBufferedReaderInitialize
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : read()
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
    Public Function read( _
        byVal alLength _
        )
        read = func_CmBufferedReaderRead(alLength)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readLine()
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
    Public Function readLine( _
        )
        readLine = func_CmBufferedReaderReadLine()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readAll()
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
    Public Function readAll( _
        )
        readAll = func_CmBufferedReaderReadAll()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : skip()
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
    Public Sub skip( _
        byVal alLength _
        )
        func_CmBufferedReaderRead alLength
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : skipLine()
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
    Public Sub skipLine( _
        )
        func_CmBufferedReaderReadLine
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : close()
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
    Public Sub close( _
        )
        sub_CmBufferedReaderClose
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
        Do While PoInbound.Item("atEndOfStream")=False And (PoBuffer.Item("length")-PoBuffer.Item("pointer")+1)<alLength
        '�C���o�E���h���ǂݏo���\�iatEndOfStream=False�j���o�b�t�@�̖��ǂݏo�������̒������ǂݍ��ޕ����������̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            func_CmBufferedReaderReadFile False
        Loop
        
        '�o�b�t�@����w�肵�����������o��
        Dim sRet : sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), alLength)
        
        '�|�C���^���X�V
        PoBuffer.Item("pointer") = PoBuffer.Item("pointer")+Len(sRet)
        
        '�A�E�g�o�E���h�̏����X�V
        Dim boFlg : If PoBuffer.Item("pointer")>PoBuffer.Item("length") Then boFlg = True Else boFlg = False
        If boFlg Then
        '�|�C���^���o�b�t�@�̊O�ɂ���ꍇ
            '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
            sub_CmBufferedReaderCopyInboundStateToOutbound
        Else
        '�|�C���^���o�b�t�@���ɂ���ꍇ
            Dim oArr : Set oArr = new_ArrSplit(Mid(PoBuffer.Item("buffer"), 1, PoBuffer.Item("pointer") - 1), vbLf)
            oArr.Reverse()
            With PoOutbound
                .Item("line") = oArr.length
                .Item("column") = Len(oArr(0))+1
                .Item("atEndOfStream") = PoInbound.Item("atEndOfStream") And (PoBuffer.Item("pointer") > PoBuffer.Item("length"))
                .Item("atEndOfLine") = .Item("atEndOfStream") Or (StrComp(Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), 1), vbLf, vbBinaryCompare)=0)
            End With
        End If
        
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
        Do While PoInbound.Item("atEndOfStream")=False And InStr(PoBuffer.Item("pointer"), PoBuffer.Item("buffer"), vbLf, vbBinaryCompare)=0
        '�C���o�E���h���ǂݏo���\�iatEndOfStream=False�j���|�C���^�̂���s���o�b�t�@�̍ŏI�s�̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            func_CmBufferedReaderReadFile False
        Loop
        
        '�s���ivbLf�j����������
        Dim lPosRowEnd : lPosRowEnd = InStr(PoBuffer.Item("pointer"), PoBuffer.Item("buffer"), vbLf, vbBinaryCompare)
        Dim sRet
        If lPosRowEnd=0 Then
        '�s���ivbLf�j��������Ȃ��������t�@�C���̏I�[�̏ꍇ
            '�|�C���^�ȍ~�S�Ă̕�����Ԃ�
            sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"))
            '�t�@�C���̏I�[�Ƀ|�C���^���X�V
            PoBuffer.Item("pointer") = PoBuffer.Item("length")+1
        Else
        '�s���ivbLf�j�������������t�@�C���̏I�[�łȂ��ꍇ
            '�|�C���^���玟�̉��s�����ivbLf�j�܂Łi���s�������܂܂Ȃ��j��Ԃ�
            sRet = Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), lPosRowEnd-PoBuffer.Item("pointer"))
           '�ŌオvbCr�̏ꍇ�͍폜����
           If StrComp(Right(sRet, 1), vbCr, vbBinaryCompare)=0 Then sRet = Mid(sRet, 1, Len(sRet)-1)
           '���̍s�̍s���i���������s���̈ʒu+1�j�Ƀ|�C���^���X�V
           PoBuffer.Item("pointer") = lPosRowEnd+1
        End If
        
        '�A�E�g�o�E���h�̏����X�V
        Dim boFlg : If PoBuffer.Item("pointer")>PoBuffer.Item("length") Then boFlg = True Else boFlg = False
        If boFlg Then
        '�|�C���^���o�b�t�@�̊O�ɂ���ꍇ
            '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
            sub_CmBufferedReaderCopyInboundStateToOutbound
        Else
        '�|�C���^���o�b�t�@���ɂ���ꍇ
            With PoOutbound
                '���̍s�̍s���ɍX�V����
                .Item("line") = .Item("line")+1
                .Item("column") = 1
                .Item("atEndOfStream") = False
                .Item("atEndOfLine") = (StrComp(Mid(PoBuffer.Item("buffer"), PoBuffer.Item("pointer"), 1), vbLf, vbBinaryCompare)=0)
            End With
        End If
        
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
        sub_CmBufferedReaderCopyInboundStateToOutbound
        '�|�C���^���X�V
        PoBuffer.Item("pointer") = PoBuffer.Item("length")+1
        
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
            .Item("buffer") = .Item("buffer") & sText
            .Item("length") = Len(.Item("buffer"))
        End With
        '�C���o�E���h�̏�Ԃ��擾����
        sub_CmBufferedReaderGetInboundStatus
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
        
        PoTextStream.Close
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
            Set PoInbound = new_DicWith(Array("line", .Line, "column", .Column, "atEndOfLine", .AtEndOfLine, "atEndOfStream", .AtEndOfStream))
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
            Set oOutbound = new_Dic()
            For Each sKey In Array("line", "column", "atEndOfLine", "atEndOfStream")
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
        sub_CmBufferedReaderGetInboundStatus
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        sub_CmBufferedReaderCopyInboundStateToOutbound
        '�|�C���^�̏�����
        Set PoBuffer = new_DicWith(Array("pointer", 1, "buffer", "", "length", 0))
    End Sub
    
End Class
