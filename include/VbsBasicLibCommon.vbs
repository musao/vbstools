'***************************************************************************************************
'FILENAME                    : VbsBasicLibCommon.vbs
'Overview                    : ���ʊ֐����C�u����
'Detailed Description        : �H����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************


'�I�t�B�X�S��

'***************************************************************************************************
'Function/Sub Name           : sub_CM_OfficeUnprotect()
'Overview                    : �����̕ی����������
'Detailed Description        : �G���[�͖�������
'                              �����̃p�X���[�h���w�肵�Ȃ��ꍇ�́A�Ăяo������vbNullString��ݒ肷�邱��
'Argument
'     aoOffice               : �I�t�B�X�̃C���X�^���X�A�G�N�Z���̏ꍇ�̓��[�N�u�b�N
'     asPassword             : �p�X���[�h
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_OfficeUnprotect( _
    byRef aoOffice _
    , byVal asPassword _
    )
    On Error Resume Next
    aoOffice.Unprotect(asPassword)
    If Err.Number Then
        Err.Clear
    End If
End Sub



'�G�N�Z���n

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcelSaveAs()
'Overview                    : �G�N�Z���t�@�C����ʖ��ŕۑ����ĕ���
'Detailed Description        : �H����
'Argument
'     aoWorkBook             : �G�N�Z���̃��[�N�u�b�N
'     asPath                 : �ۑ�����t�@�C���̃t���p�X
'     alFileformat           : XlFileFormat �񋓑́i�f�t�H���g��xlOpenXMLWorkbook 51 Excel�u�b�N�j
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcelSaveAs( _
    byRef aoWorkBook _
    , byVal asPath _
    , byVal alFileformat _
    )
    If Not(IsNumeric(alFileformat)) Then
        alFileformat = 51                  'xlOpenXMLWorkbook 51 Excel�u�b�N
    End If
    Call aoWorkBook.SaveAs( _
                            asPath _
                            , alFileformat _
                            , , _
                            , False _
                            , False _
                            )
    Call aoWorkBook.Close(False)
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelOpenFile()
'Overview                    : �G�N�Z���t�@�C����ǂݎ���p�^�_�C�A���O�Ȃ��ŊJ��
'Detailed Description        : �H����
'Argument
'     aoExcel                : �G�N�Z��
'     asPath                 : �G�N�Z���t�@�C���̃t���p�X
'Return Value
'     �J�����G�N�Z���̃��[�N�u�b�N
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelOpenFile( _
    byRef aoExcel _
    , byVal asPath _
    )    
    Set func_CM_ExcelOpenFile = aoExcel.Workbooks.Open( _
                                                        asPath _
                                                        , 0 _
                                                        , True _
                                                        , , , _
                                                        , True _
                                                        )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelGetTextFromAutoshape()
'Overview                    : �G�N�Z���̃I�[�g�V�F�C�v�̃e�L�X�g�����o��
'Detailed Description        : �G���[�͖�������
'Argument
'     aoAutoshape            : �I�[�g�V�F�C�v
'Return Value
'     �I�[�g�V�F�C�v�̃e�L�X�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    On Error Resume Next
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
    If Err.Number Then
        Err.Clear
    End If
End Function


'�t�@�C������n

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFile()
'Overview                    : �t�@�C�����폜����
'Detailed Description        : FileSystemObject��DeleteFile()�Ɠ���
'Argument
'     asPath                 : �폜����t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFile( _
    byVal asPath _
    ) 
    If Not func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFile = False
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile(asPath)
    func_CM_FsDeleteFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFolder()
'Overview                    : �t�H���_���폜����
'Detailed Description        : FileSystemObject��DeleteFolder()�Ɠ���
'Argument
'     asPath                 : �폜����t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFolder( _
    byVal asPath _
    ) 
    If Not func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFolder = False
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFolder(asPath)
    func_CM_FsDeleteFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsDeleteFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsDeleteFsObject()
'Overview                    : �t�@�C�����t�H���_���폜����
'Detailed Description        : func_CM_FsDeleteFile()��func_CM_FsDeleteFolder()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFsObject( _
    byVal asPath _
    )
    func_CM_FsDeleteFsObject = False
    If func_CM_FsFileExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFile(asPath)
    If func_CM_FsFolderExists(asPath) Then func_CM_FsDeleteFsObject = func_CM_FsDeleteFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFile ()
'Overview                    : �t�@�C�����R�s�[����
'Detailed Description        : FileSystemObject��CopyFile()�Ɠ���
'Argument
'     asPathFrom             : �R�s�[���t�@�C���̃t���p�X
'     asPathTo               : �R�s�[��t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFile = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").CopyFile(asPathFrom, asPathTo)
    func_CM_FsCopyFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFolder ()
'Overview                    : �t�H���_���R�s�[����
'Detailed Description        : FileSystemObject��CopyFolder()�Ɠ���
'Argument
'     asPathFrom             : �R�s�[���t�H���_�̃t���p�X
'     asPathTo               : �R�s�[��t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFolder = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").CopyFolder(asPathFrom, asPathTo)
    func_CM_FsCopyFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsCopyFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCopyFsObject()
'Overview                    : �t�@�C�����t�H���_���R�s�[����
'Detailed Description        : func_CM_FsCopyFile()��func_CM_FsCopyFolder()�ɈϏ�
'Argument
'     asPathFrom             : �R�s�[���t�@�C��/�t�H���_�̃t���p�X
'     asPathTo               : �R�s�[��̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCopyFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsCopyFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsCopyFsObject = func_CM_FsCopyFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFile ()
'Overview                    : �t�@�C�����ړ�����
'Detailed Description        : FileSystemObject��MoveFile()�Ɠ���
'Argument
'     asPathFrom             : �ړ����t�@�C���̃t���p�X
'     asPathTo               : �ړ���t�@�C���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFile( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFile = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").MoveFile(asPathFrom, asPathTo)
    func_CM_FsMoveFile = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFile = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFolder ()
'Overview                    : �t�H���_���ړ�����
'Detailed Description        : FileSystemObject��MoveFolder()�Ɠ���
'Argument
'     asPathFrom             : �ړ����t�H���_�̃t���p�X
'     asPathTo               : �ړ���t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFolder( _
    byVal asPathFrom _
    , byVal asPathTo _
    ) 
    If Not func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFolder = False
    On Error Resume Next
    Call CreateObject("Scripting.FileSystemObject").MoveFolder(asPathFrom, asPathTo)
    func_CM_FsMoveFolder = True
    If Err.Number Then
        Err.Clear
        func_CM_FsMoveFolder = False
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsMoveFsObject()
'Overview                    : �t�@�C�����t�H���_���ړ�����
'Detailed Description        : func_CM_FsMoveFile()��func_CM_FsMoveFolder()�ɈϏ�
'Argument
'     asPathFrom             : �ړ����t�@�C��/�t�H���_�̃t���p�X
'     asPathTo               : �ړ���̃t���p�X
'Return Value
'     ���� True:���� / False:���s
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsMoveFsObject( _
    byVal asPathFrom _
    , byVal asPathTo _
    )
    func_CM_FsMoveFsObject = False
    If func_CM_FsFileExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFile(asPathFrom, asPathTo)
    If func_CM_FsFolderExists(asPathFrom) Then func_CM_FsMoveFsObject = func_CM_FsMoveFolder(asPathFrom, asPathTo)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetParentFolderPath()
'Overview                    : �e�t�H���_�p�X�̎擾
'Detailed Description        : FileSystemObject��GetParentFolderName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �e�t�H���_�p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPath( _
    byVal asPath _
    ) 
    func_CM_FsGetParentFolderPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetBaseName()
'Overview                    : �t�@�C�����i�g���q�������j�̎擾
'Detailed Description        : FileSystemObject��GetBaseName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �t�@�C�����i�g���q�������j
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetBaseName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetGetExtensionName()
'Overview                    : �t�@�C���̊g���q�̎擾
'Detailed Description        : FileSystemObject��GetExtensionName()�Ɠ���
'Argument
'     asPath                 : �t�@�C���̃p�X
'Return Value
'     �t�@�C���̊g���q
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetGetExtensionName( _
    byVal asPath _
    ) 
    func_CM_FsGetGetExtensionName = CreateObject("Scripting.FileSystemObject").GetExtensionName(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsBuildPath()
'Overview                    : �t�@�C���p�X�̘A��
'Detailed Description        : FileSystemObject��BuildPath()�Ɠ���
'Argument
'     asFolderPath           : �p�X
'     asItemName             : asFolderPath�ɘA������t�H���_���܂��̓t�@�C����
'Return Value
'     �A�������t�@�C���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsBuildPath( _
    byVal asFolderPath _
    , byVal asItemName _
    ) 
    func_CM_FsBuildPath = CreateObject("Scripting.FileSystemObject").BuildPath(asFolderPath, asItemName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFileExists()
'Overview                    : �t�@�C���̑��݊m�F
'Detailed Description        : FileSystemObject��FileExists()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���݂��� / False:���݂��Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFileExists( _
    byVal asPath _
    ) 
    func_CM_FsFileExists = CreateObject("Scripting.FileSystemObject").FileExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsFolderExists()
'Overview                    : �t�H���_�̑��݊m�F
'Detailed Description        : FileSystemObject��FolderExists()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     ���� True:���݂��� / False:���݂��Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsFolderExists( _
    byVal asPath _
    ) 
    func_CM_FsFolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFile()
'Overview                    : �t�@�C���I�u�W�F�N�g�̎擾
'Detailed Description        : FileSystemObject��GetFile()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     File�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFile( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFile = CreateObject("Scripting.FileSystemObject").GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolder()
'Overview                    : �t�H���_�I�u�W�F�N�g�̎擾
'Detailed Description        : FileSystemObject��GetFolder()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Folder�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolder( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolder = CreateObject("Scripting.FileSystemObject").GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObject()
'Overview                    : �t�@�C�����t�H���_�I�u�W�F�N�g�̎擾
'Detailed Description        : func_CM_FsGetFile()��func_CM_FsGetFolder()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     File/Folder�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObject( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObject = Nothing
    If func_CM_FsFileExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFile(asPath)
    If func_CM_FsFolderExists(asPath) Then Set func_CM_FsGetFsObject = func_CM_FsGetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFiles()
'Overview                    : �w�肵���t�H���_�ȉ���Files�R���N�V�������擾����
'Detailed Description        : FileSystemObject��Folder�I�u�W�F�N�g��Files�R���N�V�����Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Files�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFiles( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFiles = CreateObject("Scripting.FileSystemObject").GetFolder(asPath).Files
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFolders()
'Overview                    : �w�肵���t�H���_�ȉ���Folders�R���N�V�������擾����
'Detailed Description        : FileSystemObject��Folder�I�u�W�F�N�g��SubFolders�R���N�V�����Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     Folders�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFolders( _
    byVal asPath _
    ) 
    Set func_CM_FsGetFolders = CreateObject("Scripting.FileSystemObject").GetFolder(asPath).SubFolders
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetFsObjects()
'Overview                    : �w�肵���t�H���_�ȉ���Files�R���N�V������Folders�R���N�V�������擾����
'Detailed Description        : func_CM_FsGetFiles()��func_CM_FsGetFolders()�ɈϏ�
'Argument
'     asPath                 : �p�X
'Return Value
'     Files�R���N�V������Folders�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetFsObjects( _
    byVal asPath _
    )
    Set func_CM_FsGetFsObjects = Nothing
    If Not func_CM_FsFolderExists(asPath) Then Exit Function
    Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
    With oTemp
        .Add "Filse", func_CM_FsGetFiles(asPath)
        .Add "Folders", func_CM_FsGetFolders(asPath)
    End With
    Set func_CM_FsGetFsObjects = oTemp
    Set oTemp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsGetTempFileName()
'Overview                    : �����_���ɐ������ꂽ�ꎞ�t�@�C���܂��̓t�H���_�[�̖��O�̎擾
'Detailed Description        : FileSystemObject��GetTempName()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     �ꎞ�t�@�C���܂��̓t�H���_�[�̖��O
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetTempFileName()
    func_CM_FsGetTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName()
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsCreateFolder()
'Overview                    : �t�H���_���쐬����
'Detailed Description        : FileSystemObject��CreateFolder()�Ɠ���
'Argument
'     asPath                 : �p�X
'Return Value
'     �쐬�����t�H���_�̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsCreateFolder( _
    byVal asPath _
    )
    func_CM_FsCreateFolder = CreateObject("Scripting.FileSystemObject").CreateFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsOpenTextFile()
'Overview                    : �t�@�C�����J��TextStream�I�u�W�F�N�g��Ԃ�
'Detailed Description        : FileSystemObject��OpenTextFile()�Ɠ���
'Argument
'     asPath                 : �p�X
'     alIomode               : ����/�o�̓��[�h 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPath�����݂��Ȃ��ꍇ�ɐV�����t�@�C�����쐬���邩�ǂ���
'     asFileFormat           : �t�@�C���̌`�� -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     TextStream�I�u�W�F�N�g
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsOpenTextFile( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal asFileFormat _
    )
    '�t�@�C�����J��
    Set func_CM_FsOpenTextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile( _
                                                              asPath _
                                                              , alIomode _
                                                              , aboCreate _
                                                              , asFileFormat _
                                                              )
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_FsWriteFile()
'Overview                    : �t�@�C���o�͂���
'Detailed Description        : �G���[�͖�������
'Argument
'     asPath                 : �o�͐�̃t���p�X
'     asCont                 : �o�͂�����e
'     �Ȃ�
'Return Value
'     �쐬�����t�H���_�̃t���p�X
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_FsWriteFile( _
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    '�t�@�C�����J���i���݂��Ȃ��ꍇ�͍쐬����j
    With func_CM_FsOpenTextFile(asPath, 2, True, -2)
'    With CreateObject("Scripting.FileSystemObject").OpenTextFile(asPath, 2, True)
        Call .WriteLine(asCont)
        Call .Close
    End With
    If Err.Number Then
        Err.Clear
    End If
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_FsIsSame()
'Overview                    : �w�肵���p�X�������t�@�C��/�t�H���_����������
'Detailed Description        : �H����
'Argument
'     asPathA                : �t�@�C��/�t�H���_�̃t���p�X
'     asPathB                : �t�@�C��/�t�H���_�̃t���p�X
'Return Value
'     ���� True:���� / False:����łȂ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsIsSame( _
    byVal asPathA _
    , byVal asPathB _
    )
    func_CM_FsIsSame = (func_CM_FsGetFsObject(asPathA) Is func_CM_FsGetFsObject(asPathB))
End Function



'�����񑀍�n

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConvOnlyAlphabet()
'Overview                    : �p�������啶���^�������ɕϊ�����
'Detailed Description        : �H����
'Argument
'     asTarget               : �ϊ����镶����
'     alConversion           : ���s����ϊ��̎�� 1:UpperCase,2:LowerCase
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConvOnlyAlphabet( _
    byVal asTarget _
    , byVal alConversion _
    )
    Dim sChar, asTargetTmp
    
    '1���������肷��
    Dim asTargetNew : asTargetNew = asTarget
    Dim lPos : lPos = 1
    Do While Len(asTargetNew) >= lPos
        '�ϊ��Ώۂ�1�������擾
        sChar = Mid(asTargetNew, lPos, 1)
        
        If func_CM_StrDetermineCharacterType(sChar, 1) Then
        '�ϊ��Ώۂ��p���̏ꍇ�̂ݕϊ�����
            asTargetTmp = ""
            
            '�ϊ��Ώۂ̕����܂ł̕�����
            If lPos > 1 Then
                asTargetTmp = Mid(asTargetNew, 1, lPos-1)
            End If
            
            '�ϊ�����1����������
            sChar = func_CM_StrConv(sChar, alConversion)
            asTargetTmp = asTargetTmp & sChar
            
            '�ϊ��Ώۂ̕����ڍs�̕����������
            If lPos < Len(asTargetNew) Then
                asTargetTmp = asTargetTmp & Mid(asTargetNew, lPos+1, Len(asTargetNew)-lPos)
            End If
            
            '�ϊ���̕�������i�[
            asTargetNew = asTargetTmp
        End If
        
        '�J�E���g�A�b�v
        lPos = lPos+1
    Loop
    
    func_CM_StrConvOnlyAlphabet = asTargetNew
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrDetermineCharacterType()
'Overview                    : ������𔻒肷��
'Detailed Description        : �H����
'Argument
'     asChar                 : ������𔻒肷�镶��
'     alType                 : ������ 1:���p��Alphabet
'Return Value
'     ���� True:���v���� / False:���v���Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrDetermineCharacterType( _
    byVal asChar _
    , byVal alType _
    )
    func_CM_StrDetermineCharacterType = False
    Select Case alType
        Case 1:
            If _
                    Asc("A") <= Asc(asChar) And Asc(asChar) <= Asc("Z") _
                Or Asc("a") <= Asc(asChar) And Asc(asChar) <= Asc("z")  _
                Then
                func_CM_StrDetermineCharacterType = True
            End If
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConv()
'Overview                    : ��������w��̂Ƃ���ϊ�����
'Detailed Description        : �H����
'Argument
'     asTarget               : �ϊ����镶����
'     alConversion           : ���s����ϊ��̎�ށi�����_��1,2�̂ݎ����j
'                                 1:�������啶���ɕϊ�
'                                 2:��������������ɕϊ�
'                                 3:��������̂��ׂĂ̒P��̍ŏ��̕�����啶���ɕϊ�
'                                 4:��������̋��� (1 �o�C�g) ���������C�h (2 �o�C�g) �����ɕϊ�
'                                 8:��������̃��C�h (2 �o�C�g) ���������� (1 �o�C�g) �����ɕϊ�
'                                16:��������̂Ђ炪�ȕ������J�^�J�i�����ɕϊ�
'                                32:��������̃J�^�J�i�������Ђ炪�ȕ����ɕϊ�
'Return Value
'     �ϊ�����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConv( _
    byVal asTarget _
    , byVal alalConversion _
    )
    Dim sChar, asTargetTmp
    func_CM_StrConv = asTarget
    Select Case alalConversion
        Case 1:
            func_CM_StrConv = UCase(asTarget)
        Case 2:
            func_CM_StrConv = LCase(asTarget)
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrLen()
'Overview                    : �S�p��2�����A���p��1�����Ƃ��ĕ�������Ԃ�
'Detailed Description        : �H����
'Argument
'     asTarget               : ������
'Return Value
'     ������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrLen( _
    byVal asTarget _
    )
    '1���������肷��
    Dim lLength : lLength = 0
    Dim lPos : lPos = 1
    Do While Len(asTarget) >= lPos
        '1�������擾
        sChar = Mid(asTarget, lPos, 1)
        
        If (Asc(sChar) And &HFF00) <> 0 Then
            lLength = lLength+2
        Else
            lLength = lLength+1
        End If
        
        '�J�E���g�A�b�v
        lPos = lPos+1
    Loop
    
    func_CM_StrLen = lLength
End Function


'���w�n

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathMin()
'Overview                    : �ŏ��l�����߂�
'Detailed Description        : �H����
'Argument
'     al1                    : ���l1
'     al2                    : ���l2
'Return Value
'     al1��al2�̒l����������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathMin( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 < al2 Then lRet = al1 Else lRet = al2
    func_CM_MathMin = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_MathMax()
'Overview                    : �ő�l�����߂�
'Detailed Description        : �H����
'Argument
'     al1                    : ���l1
'     al2                    : ���l2
'Return Value
'     al1��al2�̒l���傫����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_MathMax( _
    byVal al1 _ 
    , byVal al2 _
    )
    Dim lRet
    If al1 > al2 Then lRet = al1 Else lRet = al2
    func_CM_MathMax = lRet
End Function


'�z��n

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayGetDimensionNumber()
'Overview                    : �z��̎����������߂�
'Detailed Description        : �H����
'Argument
'     avArray                : �z��
'Return Value
'     �z��̎�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayGetDimensionNumber( _
    byRef avArray _ 
    )
   If Not IsArray(avArray) Then Exit Function
   On Error Resume Next
   Dim lNum : lNum = 0
   Dim lTemp
   Do
       lNum = lNum + 1
       lTemp = UBound(avArray, lNum)
   Loop Until Err.Number <> 0
   Err.Clear
   func_CM_ArrayGetDimensionNumber = lNum - 1
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ArrayAddItem()
'Overview                    : �z��ɗv�f��ǉ�����
'Detailed Description        : �H����
'Argument
'     avArray                : �z��
'     avItem                 : �ǉ�����v�f
'Return Value
'     �z��̎�����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ArrayAddItem( _
    byRef avArray _ 
    , byRef avItem _ 
    )
    On Error Resume Next
    If Ubound(avArray)>=0 Then
        Redim Preserve avArray(Ubound(avArray)+1)
    End If
    If Err.Number Then
        Redim avArray(0)
        Err.Clear
    End If
    Call sub_CM_TransferBetweenVariables(avItem, avArray(Ubound(avArray)))
End Sub

'���ꉽ�n����

'***************************************************************************************************
'Function/Sub Name           : func_CM_GetObjectByIdFromCollection()
'Overview                    : �R���N�V��������w�肵��ID�̃����o�[���擾����
'Detailed Description        : �G���[�͖�������
'Argument
'     aoClloection           : �R���N�V����
'     asId                   : ID
'Return Value
'     �Y�����郁���o�[
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_GetObjectByIdFromCollection( _
    byRef aoClloection _
    , byVal asId _
    )
    On Error Resume Next
    Dim oItem
    For Each oItem In aoClloection
        If oItem.Id = asId Then
            Set func_CM_GetObjectByIdFromCollection = oItem
            Exit Function
        End If
    Next
    Set func_CM_GetObjectByIdFromCollection = Nothing
    If Err.Number Then
        Err.Clear
    End If
    Set oItem = Nothing
End Function

''***************************************************************************************************
''Function/Sub Name           : func_CM_GetDateInMilliseconds()
''Overview                    : �������~���b�Ŏ擾����
''Detailed Description        : �H����
''Argument
''     adtDate                : ���t
''     adtTimer               : �^�C�}�[
''Return Value
''     yyyymmdd hh:mm:ss.nnnn�`���̓��t
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/10/12         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_CM_GetDateInMilliseconds( _
'    byVal adtDate _
'    , byVal adtTimer _
'    )
'    Dim dtNowTime        '���ݎ���
'    Dim lHour            '��
'    Dim lngMinute        '��
'    Dim lngSecond        '�b
'    Dim lngMilliSecond   '�~���b
'
'    dtNowTime = adtTimer
'    lMilliSecond = dtNowTime - Fix(dtNowTime)
'    lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
'    dtNowTime = Fix(dtNowTime)
'    lSecond = Right("0" & dtNowTime Mod 60, 2)
'    dtNowTime = dtNowTime \ 60
'    lMinute = Right("0" & dtNowTime Mod 60, 2)
'    dtNowTime = dtNowTime \ 60
'    lHour = Right("0" & dtNowTime, 2)
'
'    func_CM_GetDateInMilliseconds = adtDate & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
'End Function
'
''***************************************************************************************************
''Function/Sub Name           : func_CM_GetDateAsYYYYMMDD()
''Overview                    : ������YYYYMMDD�`���Ŏ擾����
''Detailed Description        : �H����
''Argument
''     �Ȃ�
''Return Value
''     yyyymmdd�`���̓��t
''---------------------------------------------------------------------------------------------------
''Histroy
''Date               Name                     Reason for Changes
''----------         ----------------------   -------------------------------------------------------
''2022/10/12         Y.Fujii                  First edition
''***************************************************************************************************
'Private Function func_CM_GetDateAsYYYYMMDD( _
'    byVal adtDate _
'    )
'    func_CM_GetDateAsYYYYMMDD = Replace(Left(adtDate,10), "/", "")
'End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CM_TransferBetweenVariables()
'Overview                    : �ϐ��Ԃ̍��ڈڑ�
'Detailed Description        : �ڑ������I�u�W�F�N�g���ۂ��ɂ��VBS�\���̈Ⴂ�iSet�̗L���j���z������
'                              �ڑ��悪�R���N�V�����̃����o�[�̏ꍇ�͓��삵�Ȃ�
'                              �ڑ��悪�ϐ��̏ꍇ�Ɏg�p�ł���
'Argument
'     avFrom                 : �ڑ����̕ϐ�
'     avTo                   : �ڑ���̕ϐ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_TransferBetweenVariables( _
    byRef avFrom _
    , byRef avTo _
    )
    If IsObject(avFrom) Then Set avTo = avFrom Else avTo = avFrom
End Sub

'***************************************************************************************************
'Function/Sub Name           : sub_CM_TransferToCollection()
'Overview                    : �ϐ��Ԃ̍��ڈڑ�
'Detailed Description        : �ڑ������I�u�W�F�N�g���ۂ��ɂ��VBS�\���̈Ⴂ�iSet�̗L���j���z������
'                              �ڑ��悪�R���N�V�����̏ꍇ�͓��֐����g�p����
'Argument
'     avFrom                 : �ڑ����̕ϐ�
'     aoCollection           : �ڑ���̃R���N�V����
'     asKey                  : �ڑ���̃R���N�V�����̃L�[
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_TransferToCollection( _
    byRef avFrom _
    , byRef aoCollection _
    , byVal asKey _
    )
    If IsObject(avFrom) Then Set aoCollection.Item(asKey) = avFrom Else aoCollection.Item(asKey) = avFrom
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_CompareVariables()
'Overview                    : �ϐ��̒l����������
'Detailed Description        : ��r����ϐ����I�u�W�F�N�g���ۂ��ɂ��VBS�\���̈Ⴂ�iSet�̗L���j���z������
'Argument
'     avA                    : ��r��
'     avB                    : ��r��
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_CompareVariables( _
    byRef avA _
    , byRef avB _
    )
    Dim boReturn : boReturn = False
    If IsObject(avB) Then boReturn = (avA Is avB) Else boReturn = (avA = avB)
    func_CM_CompareVariables = boReturn
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FillInTheCharacters()
'Overview                    : �����𖄂߂�
'Detailed Description        : �Ώە����̕s�������w�肵���A���C�����g�Ŏw�肵��������1�����ڂŖ��߂�
'                              �Ώە����ɕs�������Ȃ��ꍇ�́A�w�肵���������Ő؂���
'Argument
'     asTarget               : �Ώە�����
'     alWordCount            : ������
'     asToFillCharacter      : ���߂镶��
'     aboIsCutOut            : �������Ő؂���iTrue�F����/False�F���Ȃ��j
'     aboIsRightAlignment    : �A���C�����g�iTrue�F�E��/False�F���񂹁j
'Return Value
'     ���߂�������
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FillInTheCharacters( _
    byVal asTarget _
    , byVal alWordCount _
    , byVal asToFillCharacter _
    , byVal aboIsCutOut _
    , byVal aboIsRightAlignment _
    )
    
    '�؂���Ȃ��őΏە����񂪕��������傫���ꍇ�͏����𔲂���
    Dim lTargetLen : lTargetLen = Len(asTarget)
    If Not(aboIsCutOut) And lTargetLen>=alWordCount Then
        func_CM_FillInTheCharacters = asTarget
        Exit Function
    End If
    
    '���߂镶����̍쐬
    Dim sFillStrings : sFillStrings = ""
    If alWordCount-lTargetLen > 0 Then
        sFillStrings = String(alWordCount-lTargetLen , asToFillCharacter)
    End If
    
    Dim sResult
    '�A���C�����g�w��ɂ���ĕ�����𖄂߂�
    If aboIsRightAlignment Then
        sResult = sFillStrings & asTarget
    Else
        sResult = asTarget & sFillStrings
    End If
    
    '�؂���Ȃ��̏ꍇ�͏����𔲂���
    If Not(aboIsCutOut) Then
        func_CM_FillInTheCharacters = sResult
        Exit Function
    End If
    
    '�A���C�����g�w��ɂ���ĕ������؂���
    If aboIsRightAlignment Then
        sResult = Right(sResult, alWordCount)
    Else
        sResult = Left(sResult, alWordCount)
    End If
    func_CM_FillInTheCharacters = sResult
End Function
