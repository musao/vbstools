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
    Private PoFile
    Private PsTypeName
    
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
        Set PoFile = Nothing
        PsTypeName = "FolderItem2"
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
        Set PoTopics = Nothing
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
        Name = PoFile.Name
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
        ParentFolder = PoFile.Parent.Self.Path
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

End Class
