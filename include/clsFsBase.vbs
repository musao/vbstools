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
    Private PdbMostRecentReference                         '�L���b�V�����擾���ԁiTimer�֐��̒l�j
    Private PdbValidPeriod                                 '�L���b�V���L�����ԁi�b���j
    
    '�R���X�g���N�^
    Private Sub Class_Initialize()
        '������
        Set PoFso = Nothing
        PboUseCache = True
        PdbMostRecentReference = 0
        PdbValidPeriod = 1
        
        Set PoProp = CreateObject("Scripting.Dictionary")
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
        With PoProp
            If Not(.Exists(asKey)) Then Exit Function
            
            If func_FsBaseIsGetObjectValue(.Item(asKey)) Then
                Dim oObject : Set oObject = func_FsBaseGetObject()
                Select Case asKey
                Case "Attributes"                          '����
                    .Item(asKey) = oObject.Attributes
                Case "DateCreated"                         '�쐬���ꂽ���t�Ǝ���
                    .Item(asKey) = oObject.DateCreated
                Case "DateLastAccessed"                    '�Ō�ɃA�N�Z�X�������t�Ǝ���
                    .Item(asKey) = oObject.DateLastAccessed
                Case "DateLastModified"                    '�ŏI�X�V����
                    '�ŏI�X�V���͏�ɐݒ肷�邽�߁A�����ł͉������Ȃ�
                Case "Drive"                               '�t�@�C���܂��̓t�H���_�[������h���C�u�̃h���C�u����
                    Set .Item(asKey) = oObject.Drive
                Case "Name"                                '���O
                    .Item(asKey) = oObject.Name
                Case "ParentFolder"                        '�e�̃t�H���_�[�I�u�W�F�N�g
                    Set .Item(asKey) = oObject.ParentFolder
                Case "Path"                                '�p�X
                    .Item(asKey) = oObject.Path
                Case "ShortName"                           '�Z�����O(8.3 ���O�t���K��)
                    .Item(asKey) = oObject.ShortName
                Case "ShortPath"                           '�Z���p�X(8.3 ���O�t���K��)
                    .Item(asKey) = oObject.ShortPath
                Case "Size"                                '�T�C�Y�i�o�C�g�P�ʁj
                    .Item(asKey) = oObject.Size
                Case "Type"                                '���
                    .Item(asKey) = oObject.Type
                End Select
                '�ŏI�X�V���� �� �L���b�V�����擾���ԁiTimer�֐��̒l�j �̐ݒ�
                .Item("DateLastModified") = oObject.DateLastModified
                PdbMostRecentReference = Timer()
                Set oObject = Nothing
            End If
            
            '�l��ԋp
            If IsObject(.Item(asKey)) Then
                Set Prop = .Item(asKey)
            Else
                Prop = .Item(asKey)
            End If
        End With
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
        If Abs(Timer() - PdbMostRecentReference) > PdbValidPeriod Then
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
        If PoFso Is Nothing Then Set PoFso = CreateObject("Scripting.FileSystemObject")
        Set func_FsBaseGetFso = PoFso
    End Function

End Class
