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
    Private PsPath
    Private PsPathAlreadyOpened
    Private PsBuffer
    Private PoIomodeLst
    Private PsIomode              '����/�o�̓��[�h
    Private PoFileFormatLst
    Private PsFileFormat          '�t�@�C���̌`��
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        
        Set PoTextStream = Nothing
        PsPath = ""
        PsPathAlreadyOpened = ""
        PsBuffer = ""
        
        Set PoIomodeLst = CreateObject("Scripting.Dictionary")
        With PoIomodeLst
            .Add "ForReading", 1
            .Add "ForWriting", 2
            .Add "ForAppending", 8
        End With
        PsIomode = PoIomodeLst.Item(1)           '�f�t�H���g��ForAppending
        
        Set PoFileFormatLst = CreateObject("Scripting.Dictionary")
        With PoFileFormatLst
            .Add "TristateUseDefault", -2
            .Add "TristateTrue", -1
            .Add "TristateFalse", 0
        End With
        PsFileFormat = PoFileFormatLst.Item(1)   '�f�t�H���g��TristateUseDefault
    End Sub
    
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoFormatLst = Nothing
        Set PoIomodeLst = Nothing
        Set PoTextStream = Nothing
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
        
        '�e�L�X�g�X�g���[�����쐬����
        Call sub_CmBufferedWriterGetTextStream()
        
        PsBuffer = PsBuffer & vbCrLf & asContents
    End Sub
    
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmBufferedWriterGetTextStream()
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
    Private Sub sub_CmBufferedWriterGetTextStream( _
        )
        
        If Len(PsPath)=0 Then Exit Sub
        If func_CM_FsIsSame(PsPath, PsPathAlreadyOpened) Then
        
        
        If PoTextStream Is Nothing Then
        'PoTextStream���Ȃ���΍쐬����
            Dim boFileExists : boFileExists = func_CM_FsFileExists(PsPath)
            If Not boFileExists Then
            '�o�͐�t�@�C���̃p�X�����݂��Ȃ��ꍇ
                Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(PsPath)
                Dim boParentFolderExists : boParentFolderExists = func_CM_FsFolderExists(sParentFolderPath)
                If Not boParentFolderExists Then
                '�o�͐�t�@�C���̐e�t�H���_�����݂��Ȃ��ꍇ�A�t�H���_���쐬
                    Call func_CM_FsCreateFolder(sParentFolderPath)
                End If
            End If
            
            '�t�@�C�����J��
            Set PoTextStream = func_CM_FsOpenTextFile(PsPath, PoIomodeLst.Item(PsIomode) _
                                                  , True, PoFileFormatLst.Item(PsFileFormat))
            PsPathAlreadyOpened = PsPath
        End If
    End Sub

End Class
