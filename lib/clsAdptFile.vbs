'***************************************************************************************************
'FILENAME                    : clsAdptFile.vbs
'Overview                    : File�I�u�W�F�N�g�̃A�_�v�^�[�N���X
'Detailed Description        : File�I�u�W�F�N�g�Ɠ���IF��񋟂���
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsAdptFile
    '�N���X���ϐ��A�萔
    Private PoCacheInfo,PoCache,PoFile,PsTypeName
    
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
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PsTypeName = "FolderItem2"
        Set PoFile = Nothing
'        Set PoCacheInfo = new_DicWith(Array("ValidityPeriod", 3, "LastReferencedDateTime", Empty))
'        sub_AdptFileInitCache
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
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
'        Set PoCacheInfo = Nothing
'        Set PoCache = Nothing
        Set PoFile = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get DateLastModified()
    'Overview                    : �t�@�C���̍ŏI�X�V������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̍ŏI�X�V����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get DateLastModified()
        DateLastModified = PoFile.ModifyDate
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Name()
    'Overview                    : �t�@�C���̖��O��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̖��O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Name()
'        Name = PoFile.Name
        Name = new_Fso().GetFileName(PoFile.Path)
'        Name = func_AdptFileGet("Name")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ParentFolder()
    'Overview                    : �t�@�C���̐e�t�H���_�[�̃t���p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̐e�t�H���_�[�̃t���p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ParentFolder()
'        ParentFolder = PoFile.Parent.Self.Path
        ParentFolder = new_Fso().GetParentFolderName(PoFile.Path)
'        ParentFolder = func_AdptFileGet("ParentFolder")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Path()
    'Overview                    : �t�@�C���̃p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̃p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Path()
        Path = PoFile.Path
'        Path = func_AdptFileGet("Path")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Size()
    'Overview                    : �t�@�C���̃T�C�Y���o�C�g�P�ʂŕԂ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̃T�C�Y�i�o�C�g�P�ʁj
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Size()
        Size = PoFile.Size
'        Size = func_AdptFileGet("Size")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get FileType()
    'Overview                    : �t�@�C���̎�ނ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̎��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get [Type]()
        [Type] = PoFile.Type
'        [Type] = func_AdptFileGet("Type")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : setFileObject()
    'Overview                    : �t�@�C���̃I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoFile                 : FolderItem2�I�u�W�F�N�g
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setFileObject( _
        byRef aoFile _
        )
        If Not cf_isSame(PsTypeName, Typename(aoFile)) Then
            Err.Raise 438, "clsAdptFile.vbs:clsAdptFile+setFileObject()", "�I�u�W�F�N�g�ŃT�|�[�g����Ă��Ȃ��v���p�e�B�܂��̓��\�b�h�ł��B"
            Exit Function
        End If
        
        Set PoFile = aoFile
        Set setFileObject = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setFilePath()
    'Overview                    : �t�@�C���̃p�X����I�u�W�F�N�g��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���̃p�X
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/19         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setFilePath( _
        byVal asPath _
        )
        If Not new_Fso().FileExists(asPath) Then
            Err.Raise 76, "clsAdptFile.vbs:clsAdptFile+setFilePath()", "�p�X��������܂���B"
            Exit Function
        End If
        
        Set PoFile = new_ShellApp().Namespace(new_Fso().GetParentFolderName(asPath)).Items().Item(new_Fso().GetFileName(asPath))
        Set setFilePath = Me
    End Function


    
'    '***************************************************************************************************
'    'Function/Sub Name           : sub_AdptFileInitCache()
'    'Overview                    : �L���b�V��������������
'    'Detailed Description        : �H����
'    'Argument
'    '     �Ȃ�
'    'Return Value
'    '     �Ȃ�
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2024/01/13         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Public Sub sub_AdptFileInitCache( _
'        )
'        Set PoCache = new_DicWith(Array("DateLastModified", Empty, "Name", Empty, "ParentFolder", Empty, "Path", Empty, "Size", Empty, "Type", Empty))
'    End Sub
'    
'    '***************************************************************************************************
'    'Function/Sub Name           : func_AdptFileGet()
'    'Overview                    : �w�肵���v���p�e�B���擾����
'    'Detailed Description        : �H����
'    'Argument
'    '     asProp                 : �v���p�e�B���w�肷�镶����
'    'Return Value
'    '     �v���p�e�B�̓��e
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2024/01/13         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Public Function func_AdptFileGet( _
'        byVal asProp _
'        )
'        '�L���b�V�����p����
'        Dim boUseCache : boUseCache=False
'        If Not IsEmpty(PoCacheInfo.Item("LastReferencedDateTime")) Then
'        '�ŏI�Q�Ɠ�������łȂ��ꍇ
'            If new_Now().differenceFrom(PoCacheInfo.Item("LastReferencedDateTime"))<PoCacheInfo.Item("ValidityPeriod") Then
'            '�ŏI�Q�Ɠ�������L���b�V���L�����Ԃ��o�߂��Ă��Ȃ��ꍇ�A�Ώۂ̃L���b�V��������
'                If Not IsEmpty(PoCache.Item(asProp)) Then boUseCache=True
'            End If
'        End If
'
'        If boUseCache Then
'        '�L���b�V�����g���ꍇ
'            cf_bind func_AdptFileGet, PoCache.Item(asProp)
'            Exit Function
'        End If
'
'        '�L���b�V�����g�p���Ȃ��ꍇ
'        sub_AdptFileInitCache
'        Select Case asProp
'            Case "DateLastModified"
'                PoCache.Item(asProp) = vRet
'            Case "Name","ParentFolder","Path"
'                Dim sPath : sPath = PoFile.Path
'                PoCache.Item("Name") = new_Fso().GetFileName(sPath)
'                PoCache.Item("ParentFolder") = new_Fso().GetParentFolderName(sPath)
'                PoCache.Item("Path") = sPath
'            Case "Size"
'                PoCache.Item(asProp) = PoFile.Size
'            Case "Type"
'                PoCache.Item(asProp) = PoFile.Type
'        End Select
'        cf_bind func_AdptFileGet, PoCache.Item(asProp)
'        Set PoCacheInfo.Item("LastReferencedDateTime") = new_Now()
'    End Function

End Class
