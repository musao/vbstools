'***************************************************************************************************
'FILENAME                    : clsCmBufferedWriter.vbs
'Overview                    : �t�@�C���o�̓o�b�t�@�����O�����N���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/07         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmBufferedWriter
    '�N���X���ϐ��A�萔
    Private PoTextStream, PdbWriteBufferSize, PdbWriteIntervalTime, PoOutbound, PoInbound, PoBuffer
    
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
        Set PoTextStream = Nothing
        PdbWriteBufferSize = 5000                '�f�t�H���g��5000�o�C�g
        PdbWriteIntervalTime = 0                 '�f�t�H���g��0�b
        
        Dim vArr : vArr = Array("line", Empty, "column", Empty)
        Set PoOutbound = new_DicOf(vArr)
        Set PoInbound = new_DicOf(vArr)
        Set PoBuffer = new_DicOf(Array("buffer", Empty, "length", Empty, "lastWriteTime", Empty, "firstRequestTime", Empty))
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
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        this_close
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeBufferSize()
    'Overview                    : �o�̓o�b�t�@�T�C�Y��ݒ肷��
    'Detailed Description        : �o�͗v�����ɏo�̓o�b�t�@�̃T�C�Y������𒴂����ꍇ
    '                              �t�@�C���ɏo�͂���
    'Argument
    '     adbWriteBufferSize     : �o�̓o�b�t�@�T�C�Y
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeBufferSize( _
        byVal adbWriteBufferSize _
        )
        ast_argTrue cf_isNumeric(adbWriteBufferSize), TypeName(Me)&"+writeBufferSize() Let", "Invalid number."
        PdbWriteBufferSize = CDbl(adbWriteBufferSize)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get writeBufferSize()
    'Overview                    : �o�̓o�b�t�@�T�C�Y��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�̓o�b�t�@�T�C�Y
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get writeBufferSize()
        writeBufferSize = PdbWriteBufferSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeIntervalTime()
    'Overview                    : �o�͊Ԋu���ԁi�b�j��ݒ肷��
    'Detailed Description        : �o�͗v�����ɑO��o�͂��Ă���o�͊Ԋu���Ԃ𒴂����ꍇ
    '                              �o�̓o�b�t�@�̓��e���T�C�Y�����ł��t�@�C���ɏo�͂���
    'Argument
    '     adbWriteIntervalTime   : �o�͊Ԋu���ԁi�b�j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeIntervalTime( _
        byVal adbWriteIntervalTime _
        )
        ast_argTrue cf_isNumeric(adbWriteIntervalTime), TypeName(Me)&"+writeIntervalTime() Let", "Invalid number."
        PdbWriteIntervalTime = CDbl(adbWriteIntervalTime)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get writeIntervalTime()
    'Overview                    : �o�͊Ԋu���ԁi�b�j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�͊Ԋu���ԁi�b�j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get writeIntervalTime()
        writeIntervalTime = PdbWriteIntervalTime
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
    '2023/08/27         Y.Fujii                  First edition
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
    '2023/10/17         Y.Fujii                  First edition
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get column()
        column = PoOutbound.Item("column")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get currentBufferSize()
    'Overview                    : ���̃o�b�t�@�T�C�Y��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���̃o�b�t�@�T�C�Y
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get currentBufferSize()
        currentBufferSize = PoBuffer.Item("length")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get lastWriteTime()
    'Overview                    : �ŏI�t�@�C���o�͓�����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ŏI�t�@�C���o�͓���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get lastWriteTime()
        lastWriteTime = PoBuffer.Item("lastWriteTime")
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
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setTextStream( _
        byRef aoTextStream _
        )
        ast_argTrue this_isTextStream(aoTextStream), TypeName(Me)&"+setTextStream()", "Not a TextStream object."

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
        'Inbound�AOutbound���ŐV������
        this_updateStatus
        '�o�b�t�@�̏�����
        Set PoBuffer = new_DicOf(Array("buffer", "", "length", 0, "lastWriteTime", Empty, "firstRequestTime", Empty))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : write()
    'Overview                    : �w�肵���e�L�X�g���t�@�C���ɏ�������
    'Detailed Description        : this_write()�ɈϏ�����
    'Argument
    '     asCont                 : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub write( _
        byVal asCont _
        )
        this_writeBuffer asCont
        this_write
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeBlankLines()
    'Overview                    : �w�肵�����̉��s�������t�@�C���ɏ�������
    'Detailed Description        : this_writeFile()�ɈϏ�����
    'Argument
    '     alLines                : �t�@�C���ɏ������މ��s�����̐�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub writeBlankLines( _
        byVal alLines _
        )
        Dim sTmp : sTmp = ""
        Dim lIdx
        For lIdx=1 To alLines 
            sTmp = sTmp & vbNewLine
        Next
        this_writeBuffer sTmp
        this_write
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeLine()
    'Overview                    : �w�肵���e�L�X�g�Ɖ��s���t�@�C���ɏ�������
    'Detailed Description        : this_write()�ɈϏ�����
    'Argument
    '     asCont                 : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub writeLine( _
        byVal asCont _
        )
        this_writeBuffer asCont & vbNewLine
        this_write
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : newLine()
    'Overview                    : ���s�������t�@�C���ɏ�������
    'Detailed Description        : this_writeFile()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub newLine( _
        )
        this_writeBuffer vbNewLine
        this_write
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : flush()
    'Overview                    : �o�b�t�@�ɗ��߂����e���t�@�C���ɏo�͂���
    'Detailed Description        : this_writeFile()�ɈϏ�����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub flush( _
        )
        this_writeFile
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
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub close( _
        )
        this_close
    End Sub
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_write()
    'Overview                    : �t�@�C���o�͂���
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_write( _
        )
        '�t�@�C���o�͔��聕�t�@�C���o��
        If this_decideToWrite() Then this_writeFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_decideToWrite()
    'Overview                    : �t�@�C���o�͂��邩���f����
    'Detailed Description        : �ȉ��̏����Ŕ��f����
    '                              �E�o�b�t�@�̃T�C�Y���o�̓o�b�t�@�T�C�Y�𒴂���
    '                              �E�o�͓�������o�͊Ԋu���ԁi�b�j���o�߂���
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���� True:�t�@�C���ɏo�͂��� / False:�t�@�C���ɏo�͂��Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_decideToWrite( _
        )
        this_decideToWrite=False
        If PoTextStream Is Nothing Then Exit Function
        
        '�߂�l�̏�����
        Dim boRet : boRet=False

        If IsEmpty(PoBuffer.Item("firstRequestTime")) Then
        '����̏o�͓������Ȃ��ꍇ�A���񃊃N�G�X�g������ݒ肷��
            Set PoBuffer.Item("firstRequestTime") = new_Now()
        End If
        
        '�o�b�t�@�T�C�Y�̔���
        If PoBuffer.Item("length")>=PdbWriteBufferSize Then boRet=True
        
        If boRet Or PdbWriteIntervalTime<=0 Then
        '�o�b�t�@�̃T�C�Y���o�̓o�b�t�@�T�C�Y�𒴂������o�͓�������o�͊Ԋu���ԁi�b�j��0�ȉ��i���s�v�j�̏ꍇ�͊֐��𔲂���
            this_decideToWrite=boRet
            Exit Function
        End If
        
        '��r�p�����̎擾
        Dim oForComparison
        If IsEmpty(PoBuffer.Item("lastWriteTime")) Then
        '�O��̏o�͓������Ȃ��ꍇ�A���񃊃N�G�X�g�������g�p����
            Set oForComparison = PoBuffer.Item("firstRequestTime")
        Else
        '�O��̏o�͓���������ꍇ�A�ŏI�t�@�C���o�͓������g�p����
            Set oForComparison = PoBuffer.Item("lastWriteTime")
        End If
        
        '�o�͓����̔���
        If Abs(oForComparison.differenceFrom(new_Now()))>=PdbWriteIntervalTime Then boRet=True
        
        '�߂�l��Ԃ�
        this_decideToWrite=boRet
        
        Set oForComparison = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_writeFile()
    'Overview                    : �o�b�t�@�̓��e���t�@�C���ɏo�͂���
    'Detailed Description        : �H����
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
    Private Sub this_writeFile( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        '�t�@�C���ɏo��
        PoTextStream.Write PoBuffer.Item("buffer")
        'Inbound�AOutbound���ŐV������
        this_updateStatus
        With PoBuffer
            .Item("buffer") = ""                      '�o�b�t�@�̃N���A
            .Item("length") = 0                       '�o�b�t�@����0�ɂ���
            Set .Item("lastWriteTime") = new_Now()    '�o�͓������L�^
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_close()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : �o�b�t�@�̖��o�͕����o�͌�Ƀt�@�C���ڑ����N���[�Y����
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
    Private Sub this_close( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        '�o�b�t�@���c���Ă�����o�͂���
        If PoBuffer.Item("length")<>0 Then Call this_writeFile()
        '�e�L�X�g�X�g���[�����N���[�Y����
        PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_writeBuffer()
    'Overview                    : �o�b�t�@�ɏ�������
    'Detailed Description        : �H����
    'Argument
    '     asCont                 : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_writeBuffer( _
        byVal asCont _
        )
        Dim oArr

        With PoBuffer
            .Item("buffer") = .Item("buffer") & asCont
            .Item("length") = func_CM_StrLen(.Item("buffer"))
        End With

        Set oArr = new_ArrSplit(asCont, vbLf)
        oArr.reverse()
        With PoOutbound
            .Item("line") = .Item("line") + oArr.length-1
            If oArr.length=1 Then
                .Item("column") = .Item("column") + Len(oArr(0))
            Else
                .Item("column") = Len(oArr(0))+1
            End If
        End With

        Set oArr = Nothing
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_getInboundStatus( _
        )
        With PoTextStream
            '�C���o�E���h�̏�Ԃ��擾����
            Set PoInbound = new_DicOf(Array("line", .Line, "column", .Column))
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
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_copyInboundStateToOutbound( _
        )
        With PoInbound
            '�A�E�g�o�E���h�̏�ԂɃC���o�E���h�̏�Ԃ��R�s�[����
            Dim sKey, oOutbound
            Set oOutbound = new_Dic()
            For Each sKey In Array("line", "column")
                oOutbound.Add sKey, .Item(sKey)
            Next
        End With
        Set PoOutbound = oOutbound
        Set oOutbound = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_updateStatus()
    'Overview                    : Inbound�AOutbound���ŐV������
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_updateStatus( _
        )
        '�C���o�E���h�̏�Ԃ��擾����
        this_getInboundStatus
        '�C���o�E���h�̏�Ԃ��A�E�g�o�E���h�ɃR�s�[����
        this_copyInboundStateToOutbound
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
