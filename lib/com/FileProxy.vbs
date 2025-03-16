'***************************************************************************************************
'FILENAME                    : FileProxy.vbs
'Overview                    : File�I�u�W�F�N�g�̃v���L�V�N���X
'Detailed Description        : File�I�u�W�F�N�g�Ɠ���IF��񋟂���
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class FileProxy
    '�N���X���ϐ��A�萔
    Private PoFile,PsPath
    
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
        PsPath = vbNullString
        Set PoFile = Nothing
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
        Set PoFile = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get dateLastModified()
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
    Public Property Get dateLastModified()
        dateLastModified = this_getDateLastModified()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
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
    Public Property Get name()
        name = this_getName()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get parentFolder()
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
    Public Property Get parentFolder()
        parentFolder = this_getParentFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get path()
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
    Public Default Property Get path()
        path = this_getPath()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get size()
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
    Public Property Get size()
        size = this_getSize()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : �I�u�W�F�N�g�𕶎���ɕϊ�����
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �ϊ�����������
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = "<"&TypeName(Me)&">"&this_getPath()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get type()
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
    Public Property Get [type]()
        [type] = this_getType()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : of()
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
    Public Function of( _
        byVal asPath _
        )
        this_setData asPath, TypeName(Me)&"+of()"
        Set of = Me
    End Function
    

   
    '***************************************************************************************************
    'Function/Sub Name           : this_getDateLastModified()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getDateLastModified()
        this_getDateLastModified = Null
        If Not(PoFile Is Nothing) Then this_getDateLastModified = PoFile.ModifyDate
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_getName()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getName()
        this_getName = Null
        If Not(PoFile Is Nothing) Then this_getName = new_Fso().GetFileName(PsPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getParentFolder()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getParentFolder()
        this_getParentFolder = Null
        If Not(PoFile Is Nothing) Then this_getParentFolder = new_Fso().GetParentFolderName(PsPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getPath()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getPath()
        this_getPath = Null
        If Not(PoFile Is Nothing) Then this_getPath = PsPath
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getSize()
    'Overview                    : �t�@�C���̃T�C�Y��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���̃T�C�Y�i�o�C�g�P�ʁj
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getSize()
        this_getSize = Null
        If Not(PoFile Is Nothing) Then this_getSize = PoFile.Size
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getType()
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
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getType()
        this_getType = Null
        If Not(PoFile Is Nothing) Then this_getType = PoFile.Type
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : �f�[�^��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���̃p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setData( _
        byVal asPath _
        , byVal asSource _
        )
        ast_argNothing PoFile, asSource, "Because it is an immutable variable, its value cannot be changed."

        Dim oFile : Set oFile = Nothing
        On Error Resume Next
        Set oFile = new_FolderItem2Of(asPath)
        ast_argNotNothing PoFile, asSource, "invalid argument. " & cf_toString(asPath)

        If oFile Is Nothing Then Exit Sub

        this_setFile oFile, asSource
        this_setPath asPath, asSource
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setFile()
    'Overview                    : �t�@�C���I�u�W�F�N�g�iFolderItem2�j��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoFile                 : �t�@�C���I�u�W�F�N�g
    '     asSource               : �\�[�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setFile( _
        byRef aoFile _
        , byVal asSource _
        )
        ast_argsAreSame "FolderItem2", TypeName(aoFile), asSource, "This is not FolderItem2."
        Set PoFile = aoFile
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setPath()
    'Overview                    : �t�@�C���̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���̃p�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setPath( _
        byVal asPath _
        , byVal asSource _
        )
        PsPath = asPath
    End Sub

End Class
