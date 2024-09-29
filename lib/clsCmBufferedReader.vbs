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
        this_close
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
        ast_argTrue cf_isPositiveInteger(alReadSize), TypeName(Me)&"+readSize() Let", "Not a positive integer."
        PlReadSize = CDbl(alReadSize)
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
        ast_argTrue this_isTextStream(aoTextStream), TypeName(Me)&"+setTextStream()", "Not a TextStream object."

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
        'Inbound�AOutbound�Ȃǂ̏�������������
        this_initialize
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : read()
    'Overview                    : �t�@�C������w�肵�������������ǂݍ���
    'Detailed Description        : this_read()�ɈϏ�����
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
        read = this_read(alLength)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readLine()
    'Overview                    : �t�@�C������1�s�ǂݍ���
    'Detailed Description        : this_readLine()�ɈϏ�����
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
        readLine = this_readLine()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : readAll()
    'Overview                    : �t�@�C���S�̂�ǂݍ���
    'Detailed Description        : this_readAll()�ɈϏ�����
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
        readAll = this_readAll()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : skip()
    'Overview                    : �t�@�C������w�肵�������������X�L�b�v����
    'Detailed Description        : this_read()�ɈϏ�����
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
        this_read alLength
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : skipLine()
    'Overview                    : �t�@�C������1�s�X�L�b�v����
    'Detailed Description        : this_readLine()�ɈϏ�����
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
        this_readLine
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : close()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : this_close()�ɈϏ�����
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
        this_close
    End Sub
    
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_read()
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
    Private Function this_read( _
        byVal alLength _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�o�b�t�@���Ȃ���΃t�@�C����ǂݎ��
        Do While PoInbound.Item("atEndOfStream")=False And (PoBuffer.Item("length")-PoBuffer.Item("pointer")+1)<alLength
        '�C���o�E���h���ǂݏo���\�iatEndOfStream=False�j���o�b�t�@�̖��ǂݏo�������̒������ǂݍ��ޕ����������̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            this_readFile False
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
            this_copyInboundStateToOutbound
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
        this_read = sRet
        Set oArr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readLine()
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
    Private Function this_readLine( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�o�b�t�@���Ȃ���΃t�@�C����ǂݎ��
        Do While PoInbound.Item("atEndOfStream")=False And InStr(PoBuffer.Item("pointer"), PoBuffer.Item("buffer"), vbLf, vbBinaryCompare)=0
        '�C���o�E���h���ǂݏo���\�iatEndOfStream=False�j���|�C���^�̂���s���o�b�t�@�̍ŏI�s�̏ꍇ
            '�Ǎ��o�b�t�@�T�C�Y�����ǂݎ��
            this_readFile False
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
            this_copyInboundStateToOutbound
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
        this_readLine = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readAll()
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
    Private Function this_readAll( _
        )
        If PoTextStream Is Nothing Then Exit Function
        
        '�t�@�C���S�̂�ǂݎ��
        Dim sRet : sRet = this_readFile(True)
        
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        this_copyInboundStateToOutbound
        '�|�C���^���X�V
        PoBuffer.Item("pointer") = PoBuffer.Item("length")+1
        
        '�߂�l��Ԃ�
        this_readAll = sRet
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_readFile()
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
    Private Function this_readFile( _
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
        this_getInboundStatus
        '�߂�l��Ԃ�
        this_readFile = sText
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_close()
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
    Private Sub this_close( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getInboundStatus()
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
    Private Sub this_getInboundStatus( _
        )
        With PoTextStream
            '�C���o�E���h�̏�Ԃ��擾����
            Set PoInbound = new_DicWith(Array("line", .Line, "column", .Column, "atEndOfLine", .AtEndOfLine, "atEndOfStream", .AtEndOfStream))
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_copyInboundStateToOutbound()
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
    Private Sub this_copyInboundStateToOutbound( _
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
    'Function/Sub Name           : this_initialize()
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
    Private Sub this_initialize( _
        )
        '�C���o�E���h�̏�Ԃ��擾����
        this_getInboundStatus
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        this_copyInboundStateToOutbound
        '�|�C���^�̏�����
        Set PoBuffer = new_DicWith(Array("pointer", 1, "buffer", "", "length", 0))
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_isTextStream()
    'Overview                    : �I�u�W�F�N�g��TextStream����������
    'Detailed Description        : �H����
    'Argument
    '     aoObj                  : �I�u�W�F�N�g
    'Return Value
    '     ���� True:TextStream�ł��� / False:TextStream�łȂ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isTextStream( _
        byRef aoObj _
        )
        this_isTextStream = _
                cf_isSame(Vartype(aoObj),vbObject) _
                And _
                cf_isSame(Typename(aoObj),Typename(new_Ts(WScript.ScriptFullName,1,False,-2)))
    End Function
    
End Class
