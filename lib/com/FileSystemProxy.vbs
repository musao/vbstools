'***************************************************************************************************
'FILENAME                    : FileSystemProxy.vbs
'Overview                    : File/Folder�I�u�W�F�N�g�̃v���L�V�N���X
'Detailed Description        : File/Folder�I�u�W�F�N�g�Ɠ���IF��񋟂���
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Class FileSystemProxy
    '�N���X���ϐ��A�萔
    Private PoFolderItem,PoParent,PsPath
    
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
        Set PoFolderItem = Nothing
        Set PoParent = Nothing
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
        Set PoFolderItem = Nothing
        Set PoParent = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get allItems()
    'Overview                    : �t�H���_�[���̃A�C�e���̔z���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���N���X�̃C���X�^���X�̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get allItems( _
        )
        allItems = this_items(True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get basename()
    'Overview                    : �t�@�C���^�t�H���_�̊g���q�����������O��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̊g���q�����������O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get basename()
        basename = this_basename()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get dateLastModified()
    'Overview                    : �t�@�C���^�t�H���_�̍ŏI�X�V������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̍ŏI�X�V����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get dateLastModified()
        dateLastModified = this_dateLastModified()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get extension()
    'Overview                    : �t�@�C���^�t�H���_�̊g���q��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̊g���q
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get extension()
        extension = this_extension()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get hasItem()
    'Overview                    : �z���ɃA�C�e����1�ȏ㎝���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z���ɃA�C�e����1�ȏ㎝��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get hasItem()
        hasItem = this_hasItem()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isBrowsable()
    'Overview                    : �u���E�U�[�܂���Windows�G�N�X�v���[���[�t���[�����ŃA�C�e�����z�X�g�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     �u���E�U�[�܂���Windows�G�N�X�v���[���[�t���[�����ŃA�C�e�����z�X�g�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isBrowsable()
        isBrowsable = this_isBrowsable()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isFileSystem()
    'Overview                    : ���ڂ��t�@�C���V�X�e���̈ꕔ�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���ڂ��t�@�C���V�X�e���̈ꕔ�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isFileSystem()
        isFileSystem = this_isFileSystem()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isFolder()
    'Overview                    : �A�C�e�����t�H���_�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     �A�C�e�����t�H���_�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isFolder()
        isFolder = this_isFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get isLink()
    'Overview                    : ���ڂ��V���[�g�J�b�g�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���ڂ��V���[�g�J�b�g�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get isLink()
        isLink = this_isLink()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get items()
    'Overview                    : �t�H���_�[���̃A�C�e���̔z���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���N���X�̃C���X�^���X�̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get items( _
        )
        items = this_items(False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
    'Overview                    : �t�@�C���^�t�H���_�̖��O��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̖��O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get name()
        name = this_name()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get parentFolder()
    'Overview                    : �t�@�C���^�t�H���_�̐e�t�H���_�̃t���p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̐e�t�H���_�̓��N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get parentFolder()
        cf_bind parentFolder, this_parentFolder()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get path()
    'Overview                    : �t�@�C���^�t�H���_�̃t���p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̃t���p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get path()
        path = this_path()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get size()
    'Overview                    : �t�@�C���^�t�H���_�̃T�C�Y���o�C�g�P�ʂŕԂ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̃T�C�Y�i�o�C�g�P�ʁj
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get size()
        size = this_size()
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
        toString = "<"&TypeName(Me)&">"&this_path()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get type()
    'Overview                    : �t�@�C���^�t�H���_�̎�ނ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̎��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get [type]()
        [type] = this_type()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : of()
    'Overview                    : �t�@�C���^�t�H���_�̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���^�t�H���_�̃p�X
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
    'Function/Sub Name           : setParent()
    'Overview                    : �e�t�H���_�̃C���X�^���X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoParent               : �e�t�H���_�̃C���X�^���X
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setParent( _
        byRef aoParent _
        )
        this_setParent aoParent, TypeName(Me)&"+setParent()"
        Set setParent = Me
    End Function

    
    '***************************************************************************************************
    'Function/Sub Name           : this_basename()
    'Overview                    : �t�@�C���^�t�H���_�̊g���q�����������O��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̊g���q�����������O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_basename()
        this_basename = Null
        If this_notInInitial() Then this_basename = new_Fso().GetBaseName(PsPath)
    End Function
   
    '***************************************************************************************************
    'Function/Sub Name           : this_dateLastModified()
    'Overview                    : �t�@�C���^�t�H���_�̍ŏI�X�V������Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̍ŏI�X�V����
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_dateLastModified()
        this_dateLastModified = Null
        If this_notInInitial() Then this_dateLastModified = PoFolderItem.ModifyDate
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_extension()
    'Overview                    : �t�@�C���^�t�H���_�̊g���q��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̊g���q
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_extension()
        this_extension = Null
        If this_notInInitial() Then this_extension = new_Fso().GetExtensionName(PsPath)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_hasItem()
    'Overview                    : �z���ɃA�C�e����1�ȏ㎝���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �z���ɃA�C�e����1�ȏ㎝��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/29         Y.Fujii                  First edition
    Private Function this_hasItem()
        this_hasItem = Null
        If Not this_notInInitial() Then Exit Function

        this_hasItem = False
        If this_isFolder Then this_hasItem=(PoFolderItem.GetFolder.Items.Count>0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_items()
    'Overview                    : �t�H���_�[���̃A�C�e���̔z���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     aboRecursiveFlg        : True:�ċA�������� / False:�ċA�������Ȃ�
    'Return Value
    '     ���N���X�̃C���X�^���X�̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_items( _
        byVal aboRecursiveFlg _
        )
        this_items=Null
        If Not this_notInInitial() Then Exit Function

        this_items = Array()
        If Not this_hasItem Then Exit Function
        
        If new_Fso().FolderExists(PsPath) Then
        '�t�H���_�̏ꍇ
            this_items = this_itemsForFolder(aboRecursiveFlg)
        ElseIf this_isFolder() Then
        'zip�̏ꍇ
            this_items = this_itemsForZip(aboRecursiveFlg)
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsForFolder()
    'Overview                    : �t�H���_�[���̃A�C�e���̔z���Ԃ�
    'Detailed Description        : �t�H���_�̏ꍇ
    'Argument
    '     aboRecursiveFlg        : True:�ċA�������� / False:�ċA�������Ȃ�
    'Return Value
    '     ���N���X�̃C���X�^���X�̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_itemsForFolder( _
        byVal aboRecursiveFlg _
        )
        Dim oEle,vRet()
        With new_FolderOf(PsPath)
            '�t�@�C���̎擾
            For Each oEle In .Files
                this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg
            Next
            '�t�H���_�̎擾
            For Each oEle In .SubFolders
                this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg
            Next
        End With

        this_itemsForFolder = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsForZip()
    'Overview                    : �t�H���_�[���̃A�C�e���̔z���Ԃ�
    'Detailed Description        : zip�̏ꍇ
    'Argument
    '     aboRecursiveFlg        : True:�ċA�������� / False:�ċA�������Ȃ�
    'Return Value
    '     ���N���X�̃C���X�^���X�̔z��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_itemsForZip( _
        byVal aboRecursiveFlg _
        )
        Dim oEle,vRet()
        For Each oEle In PoFolderItem.GetFolder.Items
            this_itemsGetItems vRet,oEle.Path,aboRecursiveFlg
        Next

        this_itemsForZip = vRet
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_itemsGetItems()
    'Overview                    : �A�C�e�����擾����
    'Detailed Description        : �ċA��������ꍇ�͉��ʂ̃A�C�e�����擾����
    'Argument
    '     avAr                   : �擾�����A�C�e�����i�[����z��
    '     asPath                 : �p�X
    '     aboRecursiveFlg        : True:�ċA�������� / False:�ċA�������Ȃ�
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_itemsGetItems( _
        byRef avAr _
        , byVal asPath _
        , byVal aboRecursiveFlg _
        )
        Dim oFsProx : Set oFsProx = new_FsProxyOf(asPath).setParent(Me)
        cf_push avAr, oFsProx
        If aboRecursiveFlg And Not IsNull(oFsProx.hasItem()) Then
            If oFsProx.hasItem() Then cf_pushA avAr, oFsProx.allItems()
        End If
        Set oFsProx = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_name()
    'Overview                    : �t�@�C���^�t�H���_�̖��O��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̖��O
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_name()
        this_name = Null
        If this_notInInitial() Then this_name = new_Fso().GetFileName(PsPath)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_notInInitial()
    'Overview                    : ���C���X�^���X��������ԂłȂ����ۂ���Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���C���X�^���X��������ԂłȂ����ۂ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_notInInitial()
        this_notInInitial = Not(PoFolderItem Is Nothing)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_parentFolder()
    'Overview                    : �t�@�C���^�t�H���_�̐e�t�H���_�̃t���p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̐e�t�H���_�̓��N���X�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_parentFolder()
        this_parentFolder = Null
        If Not this_notInInitial() Then Exit Function
        If PoParent Is Nothing Then Set PoParent = new_FsProxyOf(new_Fso().GetParentFolderName(PsPath))
        Set this_parentFolder = PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_path()
    'Overview                    : �t�@�C���^�t�H���_�̃t���p�X��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̃t���p�X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_path()
        this_path = Null
        If this_notInInitial() Then this_path = PsPath
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setParent()
    'Overview                    : �e�t�H���_�̃C���X�^���X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoParent               : �e�t�H���_�̃C���X�^���X
    'Return Value
    '     ���g�̃C���X�^���X
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setParent( _
        byRef aoParent _
        , byVal asSource _
        )
        ast_argNotNothing PoFolderItem, asSource, "Please set the value before setting the parent folder."
'        ast_argNothing PoParent, asSource, "Because it is an immutable variable, its parent cannot be changed."
        ast_argsAreSame TypeName(Me), TypeName(aoParent), asSource, "This is not " & TypeName(Me) &"."
        ast_argsAreSame new_Fso().GetParentFolderName(PsPath), aoParent.path, asSource, "This is not a parent folder."

        Set PoParent = aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_size()
    'Overview                    : �t�@�C���^�t�H���_�̃T�C�Y��Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̃T�C�Y�i�o�C�g�P�ʁj
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_size()
        this_size = Null
        If Not this_notInInitial() Then Exit Function

        If new_Fso().FolderExists(PsPath) Then
        '�t�H���_�̏ꍇ
            this_size = new_FolderOf(PsPath).Size
        Else
        '�t�H���_�ȊO�̏ꍇ
            this_size = PoFolderItem.Size
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_type()
    'Overview                    : �t�@�C���^�t�H���_�̎�ނ�Ԃ�
    'Detailed Description        : �H����
    'Argument
    '     �Ȃ�
    'Return Value
    '     �t�@�C���^�t�H���_�̎��
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_type()
        this_type = Null
        If this_notInInitial() Then this_type = PoFolderItem.Type
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isBrowsable()
    'Overview                    : �u���E�U�[�܂���Windows�G�N�X�v���[���[�t���[�����ŃA�C�e�����z�X�g�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     �u���E�U�[�܂���Windows�G�N�X�v���[���[�t���[�����ŃA�C�e�����z�X�g�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isBrowsable()
        this_isBrowsable = Null
        If this_notInInitial() Then this_isBrowsable = PoFolderItem.IsBrowsable
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isFileSystem()
    'Overview                    : ���ڂ��t�@�C���V�X�e���̈ꕔ�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���ڂ��t�@�C���V�X�e���̈ꕔ�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isFileSystem()
        this_isFileSystem = Null
        If this_notInInitial() Then this_isFileSystem = PoFolderItem.IsFileSystem
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isFolder()
    'Overview                    : �A�C�e�����t�H���_�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     �A�C�e�����t�H���_�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isFolder()
        this_isFolder = Null
        If this_notInInitial() Then this_isFolder = PoFolderItem.IsFolder
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_isLink()
    'Overview                    : ���ڂ��V���[�g�J�b�g�ł��邩�ǂ�����Ԃ�
    'Detailed Description        : https://learn.microsoft.com/ja-jp/windows/win32/shell/folderitem
    'Argument
    '     �Ȃ�
    'Return Value
    '     ���ڂ��V���[�g�J�b�g�ł��邩�ǂ���
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isLink()
        this_isLink = Null
        If this_notInInitial() Then this_isLink = PoFolderItem.IsLink
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : �f�[�^��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���^�t�H���_�̃p�X
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
        ast_argNothing PoFolderItem , asSource, "Because it is an immutable variable, its value cannot be changed."

        Dim oFolderItem : Set oFolderItem = Nothing
        On Error Resume Next
        Set oFolderItem = new_FolderItem2Of(asPath)
        ast_argNotNothing PoFolderItem , asSource, "invalid argument. " & cf_toString(asPath)

        If oFolderItem Is Nothing Then Exit Sub

        this_setFolderItem oFolderItem, asSource
        this_setPath asPath, asSource
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setFolderItem()
    'Overview                    : �I�u�W�F�N�g�iFolderItem2�j��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     aoFolderItem           : �I�u�W�F�N�g
    '     asSource               : �\�[�X
    'Return Value
    '     �Ȃ�
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/03/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setFolderItem( _
        byRef aoFolderItem _
        , byVal asSource _
        )
        ast_argsAreSame "FolderItem2", TypeName(aoFolderItem), asSource, "This is not FolderItem2."
        Set PoFolderItem = aoFolderItem
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setPath()
    'Overview                    : �t�@�C���^�t�H���_�̃p�X��ݒ肷��
    'Detailed Description        : �H����
    'Argument
    '     asPath                 : �t�@�C���^�t�H���_�̃p�X
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
