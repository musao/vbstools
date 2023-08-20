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
    Private PoTextStream
    Private PoWriteDateTime
    Private PoRequestFirstDateTime
    Private PsPath
    Private PsPathAlreadyOpened
    Private PlWriteBufferSize
    Private PlWriteIntervalTime
    Private PsBuffer
    Private PoIomodeLst
    Private PsIomode              '����/�o�̓��[�h
    Private PoFileFormatLst
    Private PsFileFormat          '�t�@�C���̌`��
    
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
        PsPath = ""
        PsPathAlreadyOpened = ""
        PlWriteBufferSize = 5000                 '�f�t�H���g��5000�o�C�g
        PlWriteIntervalTime = 0                  '�f�t�H���g��0�b
        Set PoWriteDateTime = Nothing
        Set PoRequestFirstDateTime = Nothing
        PsBuffer = ""
        
        Set PoIomodeLst = CreateObject("Scripting.Dictionary")
        With PoIomodeLst
            .Add "ForReading", 1
            .Add "ForWriting", 2
            .Add "ForAppending", 8
            PsIomode = .Keys()(2)                '�f�t�H���g��ForAppending
        End With
        
        Set PoFileFormatLst = CreateObject("Scripting.Dictionary")
        With PoFileFormatLst
            .Add "TristateUseDefault", -2
            .Add "TristateTrue", -1
            .Add "TristateFalse", 0
            PsFileFormat = .Keys()(0)            '�f�t�H���g��TristateUseDefault
        End With
        
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
        Set PoWriteDateTime = Nothing
        Set PoFormatLst = Nothing
        Set PoIomodeLst = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Outpath()
    'Overview                    : �o�͐�t�@�C���̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���̃p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Outpath( _
        byVal asPath _
        )
        PsPath = asPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Outpath()
    'Overview                    : �o�͐�t�@�C���̃p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�͐�t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Outpath()
        Outpath = PsPath
    End Property
    
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
        If -2147483648<=alWriteBufferSize And alWriteBufferSize<=2147483647 Then
        'Long�^�͈̔́i-2,147,483,648 �` 2,147,483,647�j�̏ꍇ
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
    '                              �ݒ�l��0�ȉ��̏ꍇ�͂��̔��f�����Ȃ�
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
        If -2147483648<=alWriteIntervalTime And alWriteIntervalTime<=2147483647 Then
        'Long�^�͈̔́i-2,147,483,648 �` 2,147,483,647�j�̏ꍇ
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
    'Function/Sub Name           : Property Let Iomode()
    'Overview                    : ����/�o�̓��[�h��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asIomode               : ����/�o�̓��[�h "ForReading","ForWriting","ForAppending"
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Iomode( _
        byVal asIomode _
        )
        If PoIomodeLst.Exists(asIomode) Then
            PsIomode = asIomode
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Iomode()
    'Overview                    : ����/�o�̓��[�h��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�͐�t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Iomode()
        Iomode = PsIomode
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let FileFormat()
    'Overview                    : �t�@�C���̌`����ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asFileFormat           : �t�@�C���̌`�� "TristateUseDefault","TristateTrue","TristateFalse"
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let FileFormat( _
        byVal asFileFormat _
        )
        If PoFileFormatLst.Exists(asFileFormat) Then
            PsFileFormat = asFileFormat
        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileFormat()
    'Overview                    : �t�@�C���̌`����Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �o�͐�t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get FileFormat()
        FileFormat = PsFileFormat
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
    'Overview                    : �ŏI�t�@�C���o�͓���
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
            LastWriteDateTime = PoWriteDateTime.DisplayFormatAs("YYYY/MM/DD hh:mm:ss.000")
        End If
    End Property
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
    'Function/Sub Name           : WriteLineContents()
    'Overview                    : 1�s�t�@�C���o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteFile()�ɈϏ�����
    'Argument
    '     asContents             : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub WriteLineContents( _
        byVal asContents _
        )
        PsBuffer = PsBuffer & asContents & vbNewLine
        Call sub_CmBufferedWriterWriteContents()
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
        byVal asContents _
        )
        '�t�@�C���o�͔��聕�t�@�C���o��
        If func_CmBufferedWriterDetermineToWrite() Then Call sub_CmBufferedWriterWriteFile()
    End Sub
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCreateTextStream()
    'Overview                    : �e�L�X�g�X�g���[�����쐬����
    'Detailed Description        : �H����
    'Argument
    '     asContents             : ���e
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/19         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmBufferedWriterCreateTextStream( _
        )
        If Not PoTextStream Is Nothing Then
            If Len(PsPath)=0 Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
            '����PoTextStream�̖��o�͕�������������ŁA�N���[�Y����
                Call sub_CmBufferedWriterCloseFile()
            End If
        End If
        
        If Len(PsPath)>0 Then
            If PoTextStream Is Nothing Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
            'PoTextStream��V�K�쐬����
                If Not func_CM_FsFileExists(PsPath) Then
                '�o�͐�t�@�C���̃p�X�����݂��Ȃ� ���� �o�͐�t�@�C���̐e�t�H���_�����݂��Ȃ��ꍇ�A�t�H���_���쐬
                    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(PsPath)
                    If Not func_CM_FsFolderExists(sParentFolderPath) Then
                        Call func_CM_FsCreateFolder(sParentFolderPath)
                    End If
                End If
                
                'PoTextStream���쐬�i�t�@�C�����Ȃ���ΐV�K�쐬�j
                Set PoTextStream = func_CM_FsOpenTextFile(PsPath, PoIomodeLst.Item(PsIomode) _
                                                      , True, PoFileFormatLst.Item(PsFileFormat))
                '�o�͐�t�@�C���p�X��ޔ�����
                PsPathAlreadyOpened = PsPath
            End If
        End If
        
    End Sub
    
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
        '�e�L�X�g�X�g���[�����쐬����
        Call sub_CmBufferedWriterCreateTextStream()
        
        If PoTextStream Is Nothing Then Exit Sub
        
        '�t�@�C���ɏo��
        Call PoTextStream.Write(PsBuffer)
        '�o�b�t�@�̃N���A
        PsBuffer = ""
        '�o�͓������L�^
        Set PoWriteDateTime = new_clsCmCalendar()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterCloseFile()
    'Overview                    : �t�@�C���o�͂���������
    'Detailed Description        : �o�b�t�@�̖��o�͕����o�͂̏�A�t�@�C���ڑ����N���[�Y����
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
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmBufferedWriterDetermineToWrite()
    'Overview                    : �t�@�C���o�͂��邩���f����
    'Detailed Description        : �ȉ��̏����Ŕ��f����
    '                              �E�o�b�t�@�̃T�C�Y���o�̓o�b�t�@�T�C�Y�𒴂���
    '                              �E�o�͓�������o�͊Ԋu���ԁi�b�j���o�߂���
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
    Private Function func_CmBufferedWriterDetermineToWrite( _
        )
        func_CmBufferedWriterDetermineToWrite=False
        If PoTextStream Is Nothing Then Exit Function
        
        Dim boReturn : boReturn=False
        
        '�o�b�t�@�T�C�Y
        If func_CM_StrLen(PsBuffer)>=PlWriteBufferSize Then boReturn=True
        
        '�o�͓���
        If PoWriteDateTime Is Nothing Then
        '���񏑂����ݑO�͏��񃊃N�G�X�g������̌o�ߎ��ԂŔ��f����
            If PoRequestFirstDateTime Is Nothing Then
                Set PoRequestFirstDateTime = new_clsCmCalendar()
            Else
                If Abs(PoRequestFirstDateTime.DifferenceInScondsFrom(new_clsCmCalendar()))>=PlWriteIntervalTime Then boReturn=True
            End If
        Else
            If Abs(PoWriteDateTime.DifferenceInScondsFrom(new_clsCmCalendar()))>=PlWriteIntervalTime Then boReturn=True
        End If
        func_CmBufferedWriterDetermineToWrite=boReturn
    End Function

End Class
