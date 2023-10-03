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

'***************************************************************************************************
'Function/Sub Name           : new_clsCmBufferedWriter()
'Overview                    : �C���X�^���X�����֐�
'Detailed Description        : �H����
'Argument
'     aoTextStream           : �e�L�X�g�X�g���[���I�u�W�F�N�g
'Return Value
'     �t�@�C���o�̓o�b�t�@�����O�����N���X�̃C���X�^���X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmBufferedWriter( _
    byRef aoTextStream _
    )
    Set new_clsCmBufferedWriter = (New clsCmBufferedWriter).SetTextStream(aoTextStream)
End Function

Class clsCmBufferedWriter
    '�N���X���ϐ��A�萔
    Private PoTextStream
    Private PoWriteDateTime
    Private PoRequestFirstDateTime
    Private PlWriteBufferSize
    Private PlWriteIntervalTime
    Private PsBuffer
    
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
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
        PsBuffer = ""
        
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
        Call sub_CmBufferedWriterCloseFile()
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let WriteBufferSize()
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
    Public Property Let WriteBufferSize( _
        byVal alWriteBufferSize _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteBufferSize, 2) Then
            PlWriteBufferSize = CLng(alWriteBufferSize)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get WriteBufferSize()
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
    Public Property Get WriteBufferSize()
        WriteBufferSize = PlWriteBufferSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let WriteIntervalTime()
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
    Public Property Let WriteIntervalTime( _
        byVal alWriteIntervalTime _
        )
        If func_CM_ValidationlIsWithinTheRangeOf(alWriteIntervalTime, 2) Then
            PlWriteIntervalTime = CLng(alWriteIntervalTime)
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get WriteIntervalTime()
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
    Public Property Get WriteIntervalTime()
        WriteIntervalTime = PlWriteIntervalTime
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
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get TextStream()
        Set TextStream = aoTextStream
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CurrentBufferSize()
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
    Public Property Get CurrentBufferSize()
        CurrentBufferSize = func_CM_StrLen(PsBuffer)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get LastWriteDateTime()
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
    Public Property Get LastWriteDateTime()
        If PoWriteDateTime Is Nothing Then
            LastWriteDateTime=""
        Else
            LastWriteDateTime = PoWriteDateTime
        End If
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
    '2023/08/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetTextStream( _
        byRef aoTextStream _
        )
        Set PoTextStream = aoTextStream
        Set SetTextStream = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : WriteContents()
    'Overview                    : �t�@�C���o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
    'Argument
    '     asContents             : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub WriteContents( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents
        Call sub_CmBufferedWriterWriteContents()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : newLine()
    'Overview                    : ���s�������o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
    'Argument
    '     asContents             : ���e
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
        PsBuffer = PsBuffer & vbNewLine
        Call sub_CmBufferedWriterWriteContents()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Flush()
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
    Public Sub Flush( _
        )
        Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : FileClose()
    'Overview                    : �t�@�C���ڑ����N���[�Y����
    'Detailed Description        : sub_CmBufferedWriterCloseFile()�ɈϏ�����
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
    Public Sub FileClose( _
        )
        Call sub_CmBufferedWriterCloseFile()
    End Sub
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteContents()
    'Overview                    : �t�@�C���o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteContents()�ɈϏ�����
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
    Private Sub sub_CmBufferedWriterWriteContents( _
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
        Dim boReturn : boReturn=False
        
        '�o�b�t�@�T�C�Y�̔���
        If func_CM_StrLen(PsBuffer)>=PlWriteBufferSize Then boReturn=True
        
        If boReturn Or PlWriteIntervalTime<=0 Then
        '�o�b�t�@�̃T�C�Y���o�̓o�b�t�@�T�C�Y�𒴂������o�͓�������o�͊Ԋu���ԁi�b�j��0�ȉ��i���s�v�j�̏ꍇ�͊֐��𔲂���
            func_CmBufferedWriterDetermineToWrite=boReturn
            Exit Function
        End If
        
        If PoWriteDateTime Is Nothing And PoRequestFirstDateTime Is Nothing Then
        '�O��Ə���̏o�͓������Ȃ��ꍇ�A�{���N�G�X�g�i�����񃊃N�G�X�g�j�������擾���Ċ֐��𔲂���
            Set PoRequestFirstDateTime = new_clsCalGetNow()
            func_CmBufferedWriterDetermineToWrite=boReturn
            Exit Function
        End If
        
        '��r�p�����̎擾
        Dim oForComparison
        Set oForComparison = PoWriteDateTime
        If oForComparison Is Nothing Then
        '�O��̏o�͓������Ȃ��ꍇ�A���񃊃N�G�X�g�������g�p����
            Set oForComparison = PoRequestFirstDateTime
        End If
        
        '�o�͓����̔���
        If Abs(oForComparison.DifferenceInScondsFrom(new_clsCalGetNow()))>=PlWriteIntervalTime Then boReturn=True
        
        '�߂�l��Ԃ�
        func_CmBufferedWriterDetermineToWrite=boReturn
        
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
        Call PoTextStream.Write(PsBuffer)
        '�o�b�t�@�̃N���A
        PsBuffer = ""
        '�o�͓������L�^
        Set PoWriteDateTime = new_clsCalGetNow()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCloseFile()
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
    Private Sub sub_CmBufferedWriterCloseFile( _
        )
        If PoTextStream Is Nothing Then Exit Sub
        
        '�o�b�t�@���c���Ă�����o�͂���
        If func_CM_StrLen(PsBuffer)<>0 Then Call sub_CmBufferedWriterWriteFile()
        '�e�L�X�g�X�g���[�����N���[�Y����
        Call PoTextStream.Close
        Set PoTextStream = Nothing
    End Sub

End Class
