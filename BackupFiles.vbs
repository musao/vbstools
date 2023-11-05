'***************************************************************************************************
'FILENAME                    : BackupFiles.vbs
'Overview                    : �����Ŏ󂯎�����t�@�C�����o�b�N�A�b�v����
'Detailed Description        : Sendto����g�p����
'                              �t�H���_���w�肵���ꍇ�͂��̃t�H���_�ȉ��S�Ẵt�@�C�����o�b�N�A�b�v����
'Argument
'     PATH1,2...             : �t�@�C���̃p�X1,2,...
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'�萔
Private Const Cs_FOLDER_LIB = "lib"

'import��`
Sub sub_import( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_LIB)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'import
Call sub_import("clsCmArray.vbs")
Call sub_import("clsCmBufferedWriter.vbs")
Call sub_import("clsCmCalendar.vbs")
Call sub_import("clsCmBroker.vbs")
Call sub_import("clsCompareExcel.vbs")
Call sub_import("libCom.vbs")


'���C���֐����s
Call Main()
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : ���C���֐�
'Detailed Description        : �H����
'Argument
'     �Ȃ�
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    
    Dim oParams : Set oParams = new_Dic()
    
    '���X�N���v�g�̈����擾
    Call sub_BackupFilesGetParameters( _
                            oParams _
                             )
    
    '�o�b�N�A�b�v����
    Call sub_BackupFilesBackup( _
                            oParams _
                             )
    
    '�I�u�W�F�N�g���J��
    Set oParams = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_BackupFilesGetParameters()
'Overview                    : ���X�N���v�g�̈����擾
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="Parameter"�Ŋi�[����
'                              �ʃp�����[�^�i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2,...)             �ʃp�����[�^�i�[�p�n�b�V���}�b�v
'                              ���������葶�݂���t�@�C���p�X�̂ݎ擾����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFilesGetParameters( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�n�b�V���}�b�v
    Dim oParameter : Set oParameter = new_Dic()
    Dim lCnt : lCnt = 0
    Dim lFileFolderKbn : Dim sParam
    For Each sParam In WScript.Arguments
        '�t�@�C�������݂���ꍇ1�A�t�H���_�����݂���ꍇ2
        lFileFolderKbn = 0
        If new_Fso().FileExists(sParam) Then lFileFolderKbn = 1
        If new_Fso().FolderExists(sParam) Then lFileFolderKbn = 2
        
        If lFileFolderKbn Then
        '�t�@�C���܂��̓t�H���_�����݂���ꍇ�p�����[�^���擾
            lCnt = lCnt + 1
            Call oParameter.Add(lCnt, func_BackupFilesGetMapParameterInfo(lFileFolderKbn, sParam))
        End If
    Next
    
    Call aoParams.Add("Parameter", oParameter)
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_BackupFilesGetMapParameterInfo()
'Overview                    : �ʃp�����[�^�i�[�p�n�b�V���}�b�v�쐬
'Detailed Description        : �ʃp�����[�^�i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "isFile"                 True:�Ώۂ��t�@�C�� / False:�Ώۂ��t�H���_
'                              "Path"                   �t���p�X
'Argument
'     alFileFolderKbn        : �t�@�C���̏ꍇ1�A�t�H���_�̏ꍇ2
'     asPath                 : �t���p�X
'Return Value
'     �ʃp�����[�^�i�[�p�n�b�V���}�b�v
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2017/04/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_BackupFilesGetMapParameterInfo( _
    byVal alFileFolderKbn _
    , byVal asPath _
    )
    Dim oTemp : Set oTemp = new_Dic()
    Dim boIsFile : boIsFile = False
    If alFileFolderKbn = 1 Then boIsFile = True
    Call oTemp.Add("isFile", boIsFile)
    Call oTemp.Add("Path", asPath)
    Set func_BackupFilesGetMapParameterInfo = oTemp
    Set oTemp = Nothing
End Function

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : sub_BackupFilesBackup()
'Overview                    : �o�b�N�A�b�v����
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFilesBackup( _
    byRef aoParams _
    )
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    Dim lKey
    For lKey=1 To oParameter.Count
    '�p�����[�^���Ƃ̃o�b�N�A�b�v����
        Call sub_BackupFileBackupDetail(aoParams, oParameter.Item(lKey))
    Next
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : sub_BackupFileBackupDetail()
'Overview                    : �p�����[�^���Ƃ̃o�b�N�A�b�v����
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'     aoParameter            : �ʃp�����[�^�i�[�p�n�b�V���}�b�v
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFileBackupDetail( _
    byRef aoParams _
    , byRef aoParameter _
    )
    
    If aoParameter.Item("isFile") Then
    '�t�@�C���̏ꍇ
        Call sub_BackupFileProcForOneFile(aoParams, aoParameter.Item("Path"))
    Else
    '�t�H���_�̏ꍇ
        Call sub_BackupFileProcForFolder(aoParams, aoParameter.Item("Path"))
    End If
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1
'Function/Sub Name           : sub_BackupFileProcForFolder()
'Overview                    : �t�H���_�̃o�b�N�A�b�v����
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'     asPath                 : �o�b�N�A�b�v�Ώۃt�H���_�̃t���p�X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFileProcForFolder( _
    byRef aoParams _
    , byVal asPath _
    )
    
    Dim oItem
    For Each oItem In func_CM_FsGetFolders(asPath)
    '�t�H���_���̃T�u�t�H���_�̏���
        Call sub_BackupFileProcForFolder(aoParams, oItem.Path)
    Next
    For Each oItem In func_CM_FsGetFiles(asPath)
    '�t�H���_���̃t�@�C���̏���
        Call sub_BackupFileProcForOneFile(aoParams, oItem.Path)
    Next
    
    Set oItem = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-2
'Function/Sub Name           : sub_BackupFileProcForOneFile()
'Overview                    : �t�@�C�����Ƃ̃o�b�N�A�b�v����
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'     asPath                 : �o�b�N�A�b�v�Ώۃt�@�C���̃t���p�X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFileProcForOneFile( _
    byRef aoParams _
    , byVal asPath _
    )
    
    '�O�o�[�W�����̃t�@�C����T���ď����擾����
    Call sub_BackupFileFindPreviousFile(aoParams, asPath)
    
    With aoParams.Item(asPath).Item("LatestHistoryInfo")
        '�ŐV�������Ȃ� �܂��� �ŏI�X�V�������s��v �̏ꍇ�̓o�b�N�A�b�v����
        Dim boDoBackup : boDoBackup = False
        If Not(.Item("Exists")) Then
            boDoBackup = True
        ElseIf .Item("DateLastModified") <> (func_CM_FsGetFile(asPath)).DateLastModified Then
            boDoBackup = True
        End If
        
        If Not(boDoBackup) Then
        '�o�b�N�A�b�v���Ȃ��ꍇ�͊֐��𔲂���
            Exit Sub
        End If
        
        '�o�b�N�A�b�v�t�@�C�����̍쐬
        Dim sNewDate : sNewDate = new_clsCmDate().DisplaytAs("YYYYMMDD")
'        Dim sNewDate : sNewDate = func_CM_GetDateAsYYYYMMDD(Now())
        Dim sNewSeq : sNewSeq = ""
        If (StrComp(sNewDate, .Item("BackupDate"), vbBinaryCompare)=0) Then
            sNewDate = .Item("BackupDate")
            sNewSeq = Cstr(.Item("Sequence")+1)
        End If
        Dim sNewFileName
        sNewFileName = func_CM_FsGetGetBaseName(asPath) & "_"& Right(sNewDate,6)
        If (Len(sNewSeq)>0) Then sNewFileName = sNewFileName & "_" & sNewSeq
        sNewFileName = sNewFileName & "." & func_CM_FsGetGetExtensionName(asPath)
        
    End With
    
    '�R�s�[���{
    Dim sNewFilePath : sNewFilePath = func_CM_FsBuildPath(aoParams.Item(asPath).Item("OutputFolderPath"), sNewFileName)
    Call func_CM_FsCopyFile(asPath, sNewFilePath)
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-2-1
'Function/Sub Name           : sub_BackupFileFindPreviousFile()
'Overview                    : �O�o�[�W�����̃t�@�C����T���ď����擾����
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v�ɉ��L���i�[����
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              asPath�̒l               �o�b�N�A�b�v�������i�[�p�n�b�V���}�b�v
'                              
'                              �o�b�N�A�b�v�������i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              OutputFolderPath         �o�b�N�A�b�v�o�͐�
'                              LatestHistoryInfo        �ŐV�o�b�N�A�b�v�����i�[�p�n�b�V���}�b�v
'                              
'                              �ŐV�o�b�N�A�b�v�����i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Exists                   True:�ŐV���������� / False:�ŐV�������Ȃ�
'                              DateLastModified         �ŏI�X�V����
'                              Size                     �T�C�Y
'                              BackupDate               �����擾���iYYYYMMDD�`���j
'                              Sequence                 �����擾���������̏ꍇ�̘A�ԁi1,2,3,...�j
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'     asPath                 : �o�b�N�A�b�v�Ώۃt�@�C���̃t���p�X
'Return Value
'     �Ȃ�
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFileFindPreviousFile( _
    byRef aoParams _
    , byVal asPath _
    )
    
    '�o�b�N�A�b�v��t�H���_��T��
    Dim oFolders : Set oFolders = new_Dic()
    With oFolders
        Call .Add(1, "bak")
        Call .Add(2, "bk")
        Call .Add(3, "old")
    End With
    
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(asPath)
    Dim sTargetFolder : sTargetFolder = ""
    Dim sTemp : Dim lKey
    For Each lKey In oFolders.Keys
        sTemp = func_CM_FsBuildPath(sParentFolderPath, oFolders.Item(lKey))
        If new_Fso().FolderExists(sTemp) Then
            sTargetFolder = sTemp
            Exit For
        End If
    Next
    If (Len(sTargetFolder)=0) Then
        sTargetFolder = func_CM_FsBuildPath(sParentFolderPath, oFolders.Item(1))
        Call func_CM_FsCreateFolder(sTargetFolder)
    End If
    
    '�o�b�N�A�b�v�Ώۃt�@�C���̃t�@�C�����Ɗg���q����o�b�N�A�b�v�����t�@�C�����o�p�̐��K�\�����쐬
    Dim sBasename : sBasename = func_CM_FsGetGetBaseName(asPath)
    Dim sExtensionName : sExtensionName = func_CM_FsGetGetExtensionName(asPath)
    Dim sPattern
    sPattern = sBasename & "_" & "(20)?(\d{2}[01]\d[0123]\d)" & "((_)(\d+))?"
    If (Len(sExtensionName)) Then sPattern = sPattern & "." & sExtensionName
    
    '�o�b�N�A�b�v��t�H���_���璼�߂̃t�@�C��/�t�H���_��T��
    With New RegExp
        '������
        .Pattern = sPattern
        .IgnoreCase = True
        .Global = True
        
        Dim sTargetPath : sTargetPath = ""
        Dim sDate : sDate = "00010101"
        Dim lSeq : lSeq = 1
        Dim oItem : Dim sItemName : Dim sDateToComp : Dim sSeqToComp
        For Each oItem In func_CM_FsGetFiles(sTargetFolder)
            sItemName = oItem.Name
            If .Test(sItemName) Then
            '�o�b�N�A�b�v�����̏ꍇ
                '���O������t�A�A�ԕ������擾
                sDateToComp = .Replace(sItemName, "$2")
                If Len(sDateToComp)=6 Then sDateToComp = "20" & sDateToComp
                sSeqToComp = .Replace(sItemName, "$5")
                If (Len(sSeqToComp)=0) Then sSeqToComp = "1"
                
                If (sDateToComp > sDate) _
                    Or ((sDateToComp = sDate) And ( Clng(sSeqToComp) > lSeq )) _
                    Or (Len(sTargetPath)=0) Then
                '�ێ����Ă�������V�����ꍇ�A�ŐV�̃o�b�N�A�b�v�����Ƃ��ď����擾
                    sTargetPath = oItem.Path
                    sDate = sDateToComp
                    lSeq = Clng(sSeqToComp)
                End If
            End If
        Next
    End With
    Dim boExistsTargetFile : boExistsTargetFile = False
    If (Len(sTargetPath)>0) Then boExistsTargetFile = True
    Dim oTargetFile : Set oTargetFile = Nothing
    If (Len(sTargetPath)>0) Then Set oTargetFile = func_CM_FsGetFile(sTargetPath)
    
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v�Ɋi�[����
    Dim oTempHistory : Set oTempHistory = new_Dic()
    With oTempHistory
        Call .Add("Exists", boExistsTargetFile)
        If boExistsTargetFile Then
            Call .Add("DateLastModified", oTargetFile.DateLastModified)
            Call .Add("BackupDate", sDate)
            Call .Add("Sequence", lSeq)
        End If
    End With
    Dim oTempProc : Set oTempProc = new_Dic()
    With oTempProc
        Call .Add("OutputFolderPath", sTargetFolder)
        Call .Add("LatestHistoryInfo", oTempHistory)
    End With
    With aoParams
        Call .Add(asPath, oTempProc)
    End With
    
    Set oFolders = Nothing
    Set oItem = Nothing
    Set oTargetFile = Nothing
    Set oTempHistory = Nothing
    Set oTempProc = Nothing
End Sub
