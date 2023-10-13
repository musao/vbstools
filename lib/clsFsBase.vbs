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
    Private PoFso                                          'FileSystemObject�I�u�W�F�N�g
    Private PoProp                                         '�����i�[�p�n�b�V���}�b�v
    Private PboUseCache                                    '�L���b�V���g�p�ہi�ŐV���擾���邩�ǂ����j
    Private PdbLastCacheConfirmationTime                   '�ŏI�L���b�V���m�F���ԁiTimer�֐��̒l�j
    Private PdbLastCacheUpdateTime                         '�ŏI�L���b�V���X�V���ԁiTimer�֐��̒l�j
    Private PdbValidPeriod                                 '�L���b�V���L�����ԁi�b���j�A�ŏI�L���b�V���m�F���Ԃ���̌o�ߎ���
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '������
        Set PoFso = Nothing
        PboUseCache = True
        PdbLastCacheConfirmationTime = 0
        PdbLastCacheUpdateTime = 0
        PdbValidPeriod = 1
        
        Set PoProp = new_Dic()
        With PoProp
            .Add "Attributes", vbNullString                '����
            .Add "DateCreated", vbNullString               '�쐬���ꂽ���t�Ǝ���
            .Add "DateLastAccessed", vbNullString          '�Ō�ɃA�N�Z�X�������t�Ǝ���
            .Add "DateLastModified", vbNullString          '�ŏI�X�V����
            .Add "Drive", vbNullString                     '�t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
            .Add "Name", vbNullString                      '���O
            .Add "ParentFolder", vbNullString              '�e�̃t�H���_�[�I�u�W�F�N�g
            .Add "Path", vbNullString                      '�p�X
            .Add "ShortName", vbNullString                 '�Z�����O(8.3 ���O�t���K��)
            .Add "ShortPath", vbNullString                 '�Z���p�X(8.3 ���O�t���K��)
            .Add "Size", vbNullString                      '�T�C�Y�i�o�C�g�P�ʁj
            .Add "Type", vbNullString                      '���
        End With
    End Sub
    '�f�X�g���N�^
    Private Sub Class_Terminate()
        Set PoFso = Nothing
        Set PoProp = Nothing
    End Sub
    
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
        PoProp.Item("Path") = asPath
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
    'Function/Sub Name           : Property Get LastCacheConfirmationTime()
    'Overview                    : �ŏI�L���b�V���m�F���ԁiTimer�֐��̒l�j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ŏI�L���b�V���m�F���ԁiTimer�֐��̒l�j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get LastCacheConfirmationTime()
       LastCacheConfirmationTime = PdbLastCacheConfirmationTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get LastCacheUpdateTime()
    'Overview                    : �ŏI�L���b�V���X�V���ԁiTimer�֐��̒l�j��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ŏI�L���b�V���X�V���ԁiTimer�֐��̒l�j
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get LastCacheUpdateTime()
       LastCacheUpdateTime = PdbLastCacheUpdateTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Prop()
    'Overview                    : File/Folder�̑�����Ԃ�
    'Detailed Description        : �����Ŏw�肵�������̒l��ԋp����
    '                               "Attributes"        ����
    '                               "DateCreated"       �쐬���ꂽ���t�Ǝ���
    '                               "DateLastAccessed"  �Ō�ɃA�N�Z�X�������t�Ǝ���
    '                               "DateLastModified"  �ŏI�X�V����
    '                               "Drive"             �t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
    '                               "Name"              ���O
    '                               "ParentFolder"      �e�̃t�H���_�[�I�u�W�F�N�g
    '                               "Path"              �p�X
    '                               "ShortName"         �Z�����O(8.3 ���O�t���K��)
    '                               "ShortPath"         �Z���p�X(8.3 ���O�t���K��)
    '                               "Size"              �T�C�Y�i�o�C�g�P�ʁj
    '                               "Type"              ���
    'Argument
    '     asKey                  : �������w�肷��L�[
    'Return Value
    '     �����Ŏw�肵�������̒l
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/11/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Prop( _
        byVal asKey _
        )
        If Not(PoProp.Exists(asKey)) Then Exit Function
        
        If func_FsBaseIsGetObjectValue(PoProp.Item(asKey)) Then
            With func_FsBaseGetObject()
                Select Case asKey
                Case "Attributes"                          '����
                    Call cf_bindAt(PoProp, asKey, .Attributes)
'                    Call sub_CM_TransferToCollection(.Attributes, PoProp, asKey)
                Case "DateCreated"                         '�쐬���ꂽ���t�Ǝ���
                    Call cf_bindAt(PoProp, asKey, .DateCreated)
'                    Call sub_CM_TransferToCollection(.DateCreated, PoProp, asKey)
                Case "DateLastAccessed"                    '�Ō�ɃA�N�Z�X�������t�Ǝ���
                    Call cf_bindAt(PoProp, asKey, .DateLastAccessed)
'                    Call sub_CM_TransferToCollection(.DateLastAccessed, PoProp, asKey)
                Case "DateLastModified"                    '�ŏI�X�V����
                    '�ŏI�X�V���͏�ɐݒ肷�邽�߁A�����ł͉������Ȃ�
                Case "Drive"                               '�t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
                    Call cf_bindAt(PoProp, asKey, .Drive)
'                    Call sub_CM_TransferToCollection(.Drive, PoProp, asKey)
                Case "Name"                                '���O
                    Call cf_bindAt(PoProp, asKey, .Name)
'                    Call sub_CM_TransferToCollection(.Name, PoProp, asKey)
                Case "ParentFolder"                        '�e�̃t�H���_�[�I�u�W�F�N�g
                    Call cf_bindAt(PoProp, asKey, .ParentFolder)
'                    Call sub_CM_TransferToCollection(.ParentFolder, PoProp, asKey)
                Case "Path"                                '�p�X
                    Call cf_bindAt(PoProp, asKey, .Path)
'                    Call sub_CM_TransferToCollection(.Path, PoProp, asKey)
                Case "ShortName"                           '�Z�����O(8.3 ���O�t���K��)
                    Call cf_bindAt(PoProp, asKey, .ShortName)
'                    Call sub_CM_TransferToCollection(.ShortName, PoProp, asKey)
                Case "ShortPath"                           '�Z���p�X(8.3 ���O�t���K��)
                    Call cf_bindAt(PoProp, asKey, .ShortPath)
'                    Call sub_CM_TransferToCollection(.ShortPath, PoProp, asKey)
                Case "Size"                                '�T�C�Y�i�o�C�g�P�ʁj
                    Call cf_bindAt(PoProp, asKey, .Size)
'                    Call sub_CM_TransferToCollection(.Size, PoProp, asKey)
                Case "Type"                                '���
                    Call cf_bindAt(PoProp, asKey, .Type)
'                    Call sub_CM_TransferToCollection(.Type, PoProp, asKey)
                End Select
                '�ŏI�X�V���� �� �ŏI�L���b�V���X�V���ԁiTimer�֐��̒l�j �̍X�V
                Call cf_bindAt(PoProp, "DateLastModified", .DateLastModified)
'                Call sub_CM_TransferToCollection(.DateLastModified, PoProp, "DateLastModified")
                PdbLastCacheUpdateTime = Timer()
            End With
        End If
        
        '�l��ԋp
        Call cf_bind(Prop, PoProp.Item(asKey))
'        Call sub_CM_TransferBetweenVariables(PoProp.Item(asKey), Prop)
    End Function
    
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
        If Abs(Timer() - PdbLastCacheConfirmationTime) > PdbValidPeriod Then
            PdbLastCacheConfirmationTime = Timer()                   '�ŏI�L���b�V���m�F���ԁiTimer�֐��̒l�j�̍X�V
            If PoProp.Item("DateLastModified") <> func_FsBaseGetObject().DateLastModified Then Exit Function
        End If
        func_FsBaseIsGetObjectValue = False
    End Function
    
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
        Dim sPath : sPath = PoProp.Item("Path")
        With func_FsBaseGetFso()
            If .FileExists(sPath) Then Set func_FsBaseGetObject = .GetFile(sPath)
            If .FolderExists(sPath) Then Set func_FsBaseGetObject = .GetFolder(sPath)
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
        If PoFso Is Nothing Then Set PoFso = new_Fso()
        Set func_FsBaseGetFso = PoFso
    End Function

End Class
