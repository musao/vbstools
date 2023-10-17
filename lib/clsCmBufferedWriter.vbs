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
    Private PoTextStream, PlWriteBufferSize, PlWriteIntervalTime, PoOutbound, PoInbound, PoBuffer
'    Private PoTextStream, PoWriteDateTime, PoRequestFirstDateTime, PlWriteBufferSize, PlWriteIntervalTime, PsBuffer
    
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
        PlWriteBufferSize = 5000                 '�f�t�H���g��5000�o�C�g
        PlWriteIntervalTime = 0                  '�f�t�H���g��0�b
'        Set PoWriteDateTime = Nothing
'        Set PoRequestFirstDateTime = Nothing
'        PsBuffer = ""
        
        Dim vArr : vArr = Array("line", Empty, "column")
        Set PoOutbound = new_DicWith(vArr)
        Set PoInbound = new_DicWith(vArr)
        Set PoBuffer = new_DicWith(Array("buffer", Empty, "length", Empty, "lastWriteTime", Empty, "firstRequestTime", Empty))
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
        sub_CmBufferedWriterClose
        Set PoOutbound = Nothing
        Set PoInbound = Nothing
        Set PoBuffer = Nothing

'        Call sub_CmBufferedWriterClose()
'        Set PoWriteDateTime = Nothing
'        Set PoRequestFirstDateTime = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeBufferSize()
    'Overview                    : �o�̓o�b�t�@�T�C�Y��ݒ肷��
    'Detailed Description        : �o�͗v�����ɏo�̓o�b�t�@�̃T�C�Y������𒴂����ꍇ
    '                              �t�@�C���ɏo�͂���
    'Argument
    '     alWriteBufferSize      : �o�̓o�b�t�@�T�C�Y
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeBufferSize( _
        byVal alWriteBufferSize _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteBufferSize, 2) Then
            PlWriteBufferSize = CLng(alWriteBufferSize)
        Else
            Err.Raise 1031, "clsCmBufferedWriter.vbs:clsCmBufferedWriter+writeBufferSize()", "�s���Ȑ����ł��B"
        End If
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
        writeBufferSize = PlWriteBufferSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let writeIntervalTime()
    'Overview                    : �o�͊Ԋu���ԁi�b�j��ݒ肷��
    'Detailed Description        : �o�͗v�����ɑO��o�͂��Ă���o�͊Ԋu���Ԃ𒴂����ꍇ
    '                              �o�̓o�b�t�@�̓��e���T�C�Y�����ł��t�@�C���ɏo�͂���
    'Argument
    '     alWriteIntervalTime    : �o�͊Ԋu���ԁi�b�j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let writeIntervalTime( _
        byVal alWriteIntervalTime _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteIntervalTime, 2) Then
            PlWriteIntervalTime = CLng(alWriteIntervalTime)
        End If
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
        writeIntervalTime = PlWriteIntervalTime
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
'        currentBufferSize = func_CM_StrLen(PsBuffer)
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
        If IsEmpty(PoBuffer.Item("lastWriteTime")) Then
'        If PoWriteDateTime Is Nothing Then
            lastWriteTime=""
        Else
            lastWriteTime = PoBuffer.Item("lastWriteTime")
'            lastWriteTime = PoWriteDateTime
        End If
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
        If Not func_CM_UtilIsTextStream(aoTextStream) Then
            Err.Raise 438, "clsCmBufferedWriter.vbs:clsCmBufferedWriter+setTextStream()", "�I�u�W�F�N�g�ŃT�|�[�g����Ă��Ȃ��v���p�e�B�܂��̓��\�b�h�ł��B"
            Exit Function
        End If

        Set PoTextStream = aoTextStream
        Set setTextStream = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : write()
    'Overview                    : �w�肵���e�L�X�g���t�@�C���ɏ�������
    'Detailed Description        : sub_CmBufferedWriterWrite()�ɈϏ�����
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
        sub_CmBufferedWriterWriteBuffer asCont
        sub_CmBufferedWriterWrite
'        PsBuffer = PsBuffer & asContents
'        Call sub_CmBufferedWriterWrite()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeBlankLines()
    'Overview                    : �w�肵�����̉��s�������t�@�C���ɏ�������
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
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
        sub_CmBufferedWriterWriteBuffer String(alLines, vbNewLine)
        sub_CmBufferedWriterWrite
'        PsBuffer = PsBuffer & String(alLines, vbNewLine)
'        Call sub_CmBufferedWriterWrite()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : writeLine()
    'Overview                    : �w�肵���e�L�X�g�Ɖ��s���t�@�C���ɏ�������
    'Detailed Description        : sub_CmBufferedWriterWrite()�ɈϏ�����
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
        sub_CmBufferedWriterWriteBuffer asCont & vbNewLine
        sub_CmBufferedWriterWrite
'        PsBuffer = PsBuffer & asContents & vbNewLine
'        Call sub_CmBufferedWriterWrite()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : newLine()
    'Overview                    : ���s�������t�@�C���ɏ�������
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
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
        sub_CmBufferedWriterWriteBuffer vbNewLine
        sub_CmBufferedWriterWrite
'        PsBuffer = PsBuffer & vbNewLine
'        Call sub_CmBufferedWriterWrite()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : flush()
    'Overview                    : �o�b�t�@�ɗ��߂����e���t�@�C���ɏo�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
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
        sub_CmBufferedWriterWriteFile
'        Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : close()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : sub_CmBufferedWriterClose()�ɈϏ�����
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
        Call sub_CmBufferedWriterClose()
    End Sub
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWrite()
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
    Private Sub sub_CmBufferedWriterWrite( _
        )
        '�t�@�C���o�͔��聕�t�@�C���o��
        If func_CmBufferedWriterDetermineToWrite() Then Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedWriterDetermineToWrite()
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
    Private Function func_CmBufferedWriterDetermineToWrite( _
        )
        func_CmBufferedWriterDetermineToWrite=False
        If PoTextStream Is Nothing Then Exit Function
        
        '�߂�l�̏�����
        Dim boRet : boRet=False

        If IsEmpty(PoBuffer.Item("firstRequestTime")) Then
        '����̏o�͓������Ȃ��ꍇ�A���񃊃N�G�X�g������ݒ肷��
            Set PoBuffer.Item("firstRequestTime") = new_Now()
        End If
        
        '�o�b�t�@�T�C�Y�̔���
        If PoBuffer.Item("length")>=PlWriteBufferSize Then boRet=True
'        If func_CM_StrLen(PsBuffer)>=PlWriteBufferSize Then boRet=True
        
        If boRet Or PlWriteIntervalTime<=0 Then
        '�o�b�t�@�̃T�C�Y���o�̓o�b�t�@�T�C�Y�𒴂������o�͓�������o�͊Ԋu���ԁi�b�j��0�ȉ��i���s�v�j�̏ꍇ�͊֐��𔲂���
            func_CmBufferedWriterDetermineToWrite=boRet
            Exit Function
        End If
        
'        If PoWriteDateTime Is Nothing And PoRequestFirstDateTime Is Nothing Then
'        '�O��Ə���̏o�͓������Ȃ��ꍇ�A�{���N�G�X�g�i�����񃊃N�G�X�g�j�������擾���Ċ֐��𔲂���
'            Set PoRequestFirstDateTime = new_Now()
'            func_CmBufferedWriterDetermineToWrite=boRet
'            Exit Function
'        End If
        
        '��r�p�����̎擾
        Dim oForComparison
        Set oForComparison = PoBuffer.Item("lastWriteTime")
        If IsEmpty(oForComparison) Then
'        Set oForComparison = PoWriteDateTime
'        If oForComparison Is Nothing Then
        '�O��̏o�͓������Ȃ��ꍇ�A���񃊃N�G�X�g�������g�p����
            Set oForComparison = PoBuffer.Item("firstRequestTime")
'            Set oForComparison = PoRequestFirstDateTime
        End If
        
        '�o�͓����̔���
        If Abs(oForComparison.differenceFrom(new_Now()))>=PlWriteIntervalTime Then boRet=True
        
        '�߂�l��Ԃ�
        func_CmBufferedWriterDetermineToWrite=boRet
        
        Set oForComparison = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteFile()
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
    Private Sub sub_CmBufferedWriterWriteFile( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        '�t�@�C���ɏo��
        PoTextStream.Write PoBuffer.Item("buffer")
'        Call PoTextStream.Write(PsBuffer)
        '�o�b�t�@�̃N���A
        PoBuffer.Item("buffer") = ""
'        PsBuffer = ""
        '�o�͓������L�^
        Set PoBuffer.Item("lastWriteTime") = new_Now()
'        Set PoWriteDateTime = new_Now()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterClose()
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
    Private Sub sub_CmBufferedWriterClose( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        '�o�b�t�@���c���Ă�����o�͂���
        If PoBuffer.Item("length")<>0 Then Call sub_CmBufferedWriterWriteFile()
'        If func_CM_StrLen(PsBuffer)<>0 Then Call sub_CmBufferedWriterWriteFile()
        '�e�L�X�g�X�g���[�����N���[�Y����
        Call PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteBuffer()
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
    Private Sub sub_CmBufferedWriterWriteBuffer( _
        byVal asCont _
        )
        PoBuffer.Item("buffer") = PoBuffer.Item("buffer") & asCont
        PoBuffer.Item("length") = func_CM_StrLen(PoBuffer.Item("buffer"))
    End Sub

End Class
