'***************************************************************************************************
'FILENAME                    : BackupFiles.vbs
'Overview                    : �����Ŏ󂯎�����t�@�C�����o�b�N�A�b�v����
'Detailed Description        : Sendto����g�p����
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
Private Const Cs_FOLDER_INCLUDE = "include"

'Include�p�֐���`
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'Include
Call sub_Include("VbsBasicLibCommon.vbs")


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
    
    Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
    
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
    Dim oParameter : Set oParameter = CreateObject("Scripting.Dictionary")
    Dim lCnt : lCnt = 0
    Dim lFileFolderKbn : Dim sParam
    For Each sParam In WScript.Arguments
        '�t�@�C�������݂���ꍇ1�A�t�H���_�����݂���ꍇ2
        lFileFolderKbn = 0
        If func_CM_FsFileExists(sParam) Then lFileFolderKbn = 1
        If func_CM_FsFolderExists(sParam) Then lFileFolderKbn = 2
        
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
    Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
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
    '�t�@�C�����Ƃɏ�������
        Call sub_BackupFileBackupDetail(aoParams, oParameter.Item(lKey))
    Next
    
    '�I�u�W�F�N�g���J��
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : sub_BackupFileBackupDetail()
'Overview                    : �t�@�C�����Ƃ̃o�b�N�A�b�v����
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
    
    '�O�o�[�W�����̃t�@�C����T���ď����擾����
    Call sub_BackupFileFindPreviousFile(aoParams, aoParameter)
    
    Call Msgbox("OutputFolderPath : " & aoParams.Item("OutputFolderPath"))
    Call Msgbox("Path : " & aoParams.Item("LatestHistoryInfo").Item("Path"))
    Call Msgbox("Date : " & aoParams.Item("LatestHistoryInfo").Item("Date"))
    Call Msgbox("Sequence : " & aoParams.Item("LatestHistoryInfo").Item("Sequence"))
    '�o�b�N�A�b�v�v�۔��f
    
    '�o�b�N�A�b�v���{
     '�o�b�N�A�b�v�t�@�C�����̊m��
     '�R�s�[���{
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1
'Function/Sub Name           : sub_BackupFileFindPreviousFile()
'Overview                    : �O�o�[�W�����̃t�@�C����T���ď����擾����
'Detailed Description        : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v�ɉ��L���i�[����
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              OutputFolderPath         �o�b�N�A�b�v�o�͐�̃p�X
'                              
'                              �p�����[�^�i�[�p�ėp�n�b�V���}�b�v��Key="LatestHistoryInfo"�Ŋi�[����
'                              �ŐV�o�b�N�A�b�v�����i�[�p�n�b�V���}�b�v�̍\��
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Path                     �p�X
'                              Date                     �����擾���iYYYYMMDD�`���j
'                              Sequence                 �A�ԁi1,2,3,...�j
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
Private Sub sub_BackupFileFindPreviousFile( _
    byRef aoParams _
    , byRef aoParameter _
    )
    
    '�o�b�N�A�b�v��t�H���_��T��
    Dim oFolders : Set oFolders = CreateObject("Scripting.Dictionary")
    With oFolders
        Call .Add(1, "bak")
        Call .Add(2, "bk")
        Call .Add(3, "old")
    End With
    
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(aoParameter.Item("Path"))
    Dim sTargetFolder : sTargetFolder = sParentFolderPath
    Dim sTemp : Dim lKey
    For Each lKey In oFolders.Keys
        sTemp = func_CM_FsBuildPath(sParentFolderPath, oFolders.Item(lKey))
        If func_CM_FsFolderExists(sTemp) Then
            sTargetFolder = sTemp
            Exit For
        End If
    Next
    
    '�o�b�N�A�b�v�Ώۃt�@�C���̃t�@�C�����Ɗg���q����o�b�N�A�b�v�����t�@�C�����o�p�̐��K�\�����쐬
    Dim sBasename : sBasename = func_CM_FsGetGetBaseName(aoParameter.Item("Path"))
    Dim sExtensionName : sExtensionName = func_CM_FsGetGetExtensionName(aoParameter.Item("Path"))
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
        For Each oItem In func_BackupFilesGetTargetList(aoParameter, sTargetFolder)
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
                    Or Not(Len(sTargetPath)) Then
                '�ێ����Ă�������V�����ꍇ�A�ŐV�̃o�b�N�A�b�v�����Ƃ��ď����擾
                    sTargetPath = oItem.Path
                    sDate = sDateToComp
                    lSeq = Clng(sSeqToComp)
                End If
            End If
        Next
    End With
    
    '�p�����[�^�i�[�p�ėp�n�b�V���}�b�v�Ɋi�[����
    Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
    With oTemp
        Call .Add("Path", sTargetPath)
        Call .Add("Date", sDate)
        Call .Add("Sequence", lSeq)
    End With
    With aoParams
        Call .Add("OutputFolderPath", sTargetFolder)
        Call .Add("LatestHistoryInfo", oTemp)
    End With
    
    Set oFolders = Nothing
    Set oItem = Nothing
    Set oTemp = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1-1
'Function/Sub Name           : func_BackupFilesGetTargetList()
'Overview                    : �o�b�N�A�b�v��t�H���_�̃t�@�C���܂��̓t�H���_�̃��X�g���擾
'Detailed Description        : �H����
'Argument
'     aoParams               : �p�����[�^�i�[�p�ėp�n�b�V���}�b�v
'     aoParameter            : �ʃp�����[�^�i�[�p�n�b�V���}�b�v
'Return Value
'     Files�܂���Folders�R���N�V����
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_BackupFilesGetTargetList( _
    byRef aoParameter _
    , byVal asTargetFolder _
    )
    Set func_BackupFilesGetTargetList = Nothing
    
    '�o�b�N�A�b�v�t�@�C�����璼�߂̃t�@�C��/�t�H���_��T��
    If aoParameter.Item("isFile") Then
        Set func_BackupFilesGetTargetList = func_CM_FsGetFiles(asTargetFolder)
    Else
        Set func_BackupFilesGetTargetList = func_CM_FsGetFolders(asTargetFolder)
    End If
    
End Function
