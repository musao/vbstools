'***************************************************************************************************
'FILENAME                    : clsFsBase.vbs
'Overview                    : �t�@�C���E�t�H���_���ʃN���X
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Class clsFsBase
    '�N���X���ϐ��A�萔
    Private PoFso                              'FileSystemObject�I�u�W�F�N�g
    Private PboUseCache                        '�L���b�V���g�p�ہi�ŐV���擾���邩�ǂ����j
    Private PdbMostRecentReference             '�L���b�V�����擾���ԁiTimer�֐��̒l�j
    Private PdbValidPeriod                     '�L���b�V���L�����ԁi�b���j
    
    Private PlAttributes                       '����
    Private PdtDateCreated                     '�쐬���ꂽ���t�Ǝ���
    Private PdtDateLastAccessed                '�Ō�ɃA�N�Z�X�������t�Ǝ���
    Private PdtDateLastModified                '�ŏI�X�V����
    Private PsDrive                            '�t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
    Private PsName                             '���O
    Private PoParentFolder                     '�e�̃t�H���_�[�I�u�W�F�N�g
    Private PsPath                             '�p�X
    Private PsShortName                        '�Z�����O(8.3 ���O�t���K��)
    Private PsShortPath                        '�Z���p�X(8.3 ���O�t���K��)
    Private PlSize                             '�T�C�Y�i�o�C�g�P�ʁj
    Private PsType                             '���
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '������
        Set PoFso = Nothing
        PboUseCache = True
        PdbMostRecentReference = 0
        PdbValidPeriod = 1
        
        PlAttributes = vbNullString
        PdtDateCreated = vbNullString
        PdtDateLastAccessed = vbNullString
        PdtDateLastModified = vbNullString
        PsDrive = vbNullString
        PsName = vbNullString
        PoParentFolder = vbNullString
        PsPath = vbNullString
        PsShortName = vbNullString
        PsShortPath = vbNullString
        PlSize = vbNullString
        PsType = vbNullString
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoFso = Nothing
        Set PoParentFolder = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Attributes()
    'Overview                    : ������Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     ����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Attributes( _
        )
        If func_FsBaseIsGetObjectValue(PlAttributes) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PlAttributes = oObject.Attributes
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Attributes = PlAttributes
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateCreated()
    'Overview                    : �쐬���ꂽ���t�Ǝ�����Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �쐬���ꂽ���t�Ǝ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateCreated( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateCreated) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PdtDateCreated = oObject.DateCreated
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateCreated = PdtDateCreated
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastAccessed()
    'Overview                    : �Ō�ɃA�N�Z�X�������t�Ǝ�����Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Ō�ɃA�N�Z�X�������t�Ǝ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastAccessed( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateLastAccessed) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PdtDateLastAccessed = oObject.DateLastAccessed
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateLastAccessed = PdtDateLastAccessed
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastModified()
    'Overview                    : �ŏI�X�V������Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ŏI�X�V����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastModified( _
        )
        If func_FsBaseIsGetObjectValue(PdtDateLastModified) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        DateLastModified = PdtDateLastModified
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Drive()
    'Overview                    : �t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u������Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Drive( _
        )
        If func_FsBaseIsGetObjectValue(PsDrive) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsDrive = oObject.Drive
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Drive = PsDrive
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Name()
    'Overview                    : ���O��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Name( _
        )
        If func_FsBaseIsGetObjectValue(PsName) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsName = oObject.Name
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Name = PsName
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ParentFolder()
    'Overview                    : �e�̃t�H���_�[�I�u�W�F�N�g��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �e�̃t�H���_�[�I�u�W�F�N�g
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ParentFolder( _
        )
        If func_FsBaseIsGetObjectValue(PoParentFolder) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            Set PoParentFolder = oObject.ParentFolder
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Set ParentFolder = PoParentFolder
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Patha()
    'Overview                    : �p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Path( _
        byVal asPath _
        )
        PsPath = asPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Path()
    'Overview                    : �p�X��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Path( _
        )
        If func_FsBaseIsGetObjectValue(PsPath) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsPath = oObject.Path
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Path = PsPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ShortName()
    'Overview                    : �Z�����O(8.3 ���O�t���K��)��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Z�����O(8.3 ���O�t���K��)
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ShortName( _
        )
        If func_FsBaseIsGetObjectValue(PsShortName) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsShortName = oObject.ShortName
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        ShortName = PsShortName
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ShortPath()
    'Overview                    : �Z���p�X(8.3 ���O�t���K��)��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �Z���p�X(8.3 ���O�t���K��)
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ShortPath( _
        )
        If func_FsBaseIsGetObjectValue(PsShortPath) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsShortPath = oObject.ShortPath
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        ShortPath = PsShortPath
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Size()
    'Overview                    : �T�C�Y�i�o�C�g�P�ʁj��Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     �T�C�Y�i�o�C�g�P�ʁj
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Size( _
        )
        If func_FsBaseIsGetObjectValue(PlSize) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PlSize = oObject.Size
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        Size = PlSize
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileFolderType()
    'Overview                    : ��ނ�Ԃ�
    'Detailed Description        : File/Folder�I�u�W�F�N�g�̓����v���p�e�B�Ɠ��l
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get FileFolderType( _
        )
        If func_FsBaseIsGetObjectValue(PsType) Then
            Dim oObject : Set oObject = func_FsBaseGetObject()
            PsType = oObject.Type
            Call sub_FsBaseRecordCacheAcquisition(oObject)
            Set oObject = Nothing
        End If
        FileFolderType = PsType
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Fso()
    'Overview                    : �{�C���X�^���X���g�p����FileSystemObject�I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoFso                  : FileSystemObject�I�u�W�F�N�g
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Fso( _
        byRef aoFso _
        )
        Set PoFso = aoFso
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let UseCache()
    'Overview                    : �L���b�V���g�p�ہi�ŐV���擾���邩�ǂ����j��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aboUseCache            : �L���b�V���g�p��
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let UseCache( _
        byVal aboUseCache _
        )
        PboUseCache = aboUseCache
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get UseCache()
    'Overview                    : �L���b�V���g�p�ہi�ŐV���擾���邩�ǂ����j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �L���b�V���g�p��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get UseCache()
       UseCache = PboUseCache
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let ValidPeriod()
    'Overview                    : �L���b�V���L�����ԁi�b���j��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     adbValidPer            : �L���b�V���L�����ԁi�b���j
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let ValidPeriod( _
        byVal adbValidPeriod _
        )
        PdbValidPeriod = adbValidPeriod
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ValidPeriod()
    'Overview                    : �L���b�V���L�����ԁi�b���j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �L���b�V���L�����ԁi�b���j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ValidPeriod()
       ValidPeriod = PdbValidPeriod
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get MostRecentReference()
    'Overview                    : �L���b�V�����擾���ԁiTimer�֐��̒l�j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �L���b�V�����擾���ԁiTimer�֐��̒l�j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get MostRecentReference()
       MostRecentReference = PdbMostRecentReference
    End Property
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseIsGetObjectValue()
    'Overview                    : File/Folder�I�u�W�F�N�g����l���擾���邩���f����
    'Detailed Description        : ���L�����ꂩ�ɊY������ꍇ�̓I�u�W�F�N�g���Q�Ƃ���
    '                              �E�L���b�V�����Ȃ��i�Q�Ƃ���l��vbNullString�j
    '                              �E��L�ȊO�ŁA�L���b�V�����g�p���Ȃ�
    '                              �E��L�ȊO�ŁA�L�����Ԃ𒴉߂����Y�I�u�W�F�N�g�̍ŏI�X�V�����ς����
    'Argument
    '     avSomeValue            : �Q�Ƃ���l
    'Return Value
    '     ���� True:�v / False:��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseIsGetObjectValue( _
        byRef avSomeValue _
        )
        func_FsBaseIsGetObjectValue = True
        If avSomeValue = vbNullString Then Exit Function
        If Not(PboUseCache) Then Exit Function
        If Abs(Timer() - PdbMostRecentReference) > PdbValidPeriod Then
            If PdtDateLastModified <> func_FsBaseGetObject().DateLastModified Then Exit Function
        End If
        func_FsBaseIsGetObjectValue = False
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_FsBaseRecordCacheAcquisition()
    'Overview                    : �L���b�V���擾���̏����L�^����
    'Detailed Description        : ���L���L�^����
    '                              �E�ŏI�X�V����
    '                              �E�L���b�V�����擾���ԁiTimer�֐��̒l�j
    'Argument
    '     aoSomeObject           : File/Folder�I�u�W�F�N�g
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_FsBaseRecordCacheAcquisition( _
        byRef aoSomeObject _
        )
        With func_FsBaseGetObject()
            PdtDateLastModified = .DateLastModified
            PdbMostRecentReference = Timer()
        End With
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseGetObject()
    'Overview                    : File/Folder�I�u�W�F�N�g���擾����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     File/Folder�I�u�W�F�N�g
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseGetObject( _
        )
        Set func_FsBaseGetObject = Nothing
        With func_FsBaseGetFso()
            If .FileExists(PsPath) Then Set func_FsBaseGetObject = .GetFile(PsPath)
            If .FolderExists(PsPath) Then Set func_FsBaseGetObject = .GetFolder(PsPath)
        End With
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_FsBaseGetFso()
    'Overview                    : FileSystemObject�I�u�W�F�N�g���擾����
    'Detailed Description        : Nothing��������쐬����
    'Argument
    '     �Ȃ�
    'Return Value
    '     FileSystemObject�I�u�W�F�N�g
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_FsBaseGetFso( _
        )
        If PoFso Is Nothing Then Set PoFso = CreateObject("Scripting.FileSystemObject")
        Set func_FsBaseGetFso = PoFso
    End Function

End Class
