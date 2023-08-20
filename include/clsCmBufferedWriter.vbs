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
        
        Dim vKeys
        
        Set PoTextStream = Nothing
        PsPath = ""
        PsPathAlreadyOpened = ""
        PlWriteBufferSize = 5000                 '�f�t�H���g��5000�o�C�g
        PlWriteIntervalTime = 60                 '�f�t�H���g��60�b
        Set PoWriteDateTime = Nothing
        PsBuffer = ""
        
        Set PoIomodeLst = CreateObject("Scripting.Dictionary")
        With PoIomodeLst
            .Add "ForReading", 1
            .Add "ForWriting", 2
            .Add "ForAppending", 8
            
            vKeys = .Keys
            PsIomode = vKeys(2)                  '�f�t�H���g��ForAppending
        End With
        
        Set PoFileFormatLst = CreateObject("Scripting.Dictionary")
        With PoFileFormatLst
            .Add "TristateUseDefault", -2
            .Add "TristateTrue", -1
            .Add "TristateFalse", 0
            
            vKeys = .Keys
            PsFileFormat = vKeys(0)              '�f�t�H���g��TristateUseDefault
        End With
        
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
        Set PoTextStream = Nothing
        Set PoFormatLst = Nothing
        Set PoIomodeLst = Nothing
        Set PoWriteDateTime = Nothing
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
    'Function/Sub Name           : WriteContents()
    'Overview                    : �t�@�C���o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteContents()�ɈϏ�����
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
        Call sub_CmBufferedWriterWriteContents(asContents)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterWriteContents()
    'Overview                    : �t�@�C���o�͂���
    'Detailed Description        : sub_CmBufferedWriterWriteContents()�ɈϏ�����
    'Argument
    '     asContents             : ���e
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
        
        PsBuffer = PsBuffer & vbCrLf & asContents
        
        '�e�L�X�g�X�g���[�����쐬����
        Call sub_CmBufferedWriterCreateTextStream()
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
        
        If Len(PsPathAlreadyOpened)>0 Then
            If Len(PsPath)=0 Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
            '����PoTextStream�̖��o�͕�������������ŁA�N���[�Y����
                '���H����
            End If
        End If
        
        If Len(PsPath)>0 Then
            If Len(PsPathAlreadyOpened)=0 Or Not func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
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
                PsPathAlreadyOpened = PsPath
            End If
        End If
        
    End Sub

End Class
