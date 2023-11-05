'***************************************************************************************************
'FILENAME                    : BackupFiles.vbs
'Overview                    : 引数で受け取ったファイルをバックアップする
'Detailed Description        : Sendtoから使用する
'                              フォルダを指定した場合はそのフォルダ以下全てのファイルをバックアップする
'Argument
'     PATH1,2...             : ファイルのパス1,2,...
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'定数
Private Const Cs_FOLDER_LIB = "lib"

'import定義
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


'メイン関数実行
Call Main()
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : メイン関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数取得
    Call sub_BackupFilesGetParameters( _
                            oParams _
                             )
    
    'バックアップする
    Call sub_BackupFilesBackup( _
                            oParams _
                             )
    
    'オブジェクトを開放
    Set oParams = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : sub_BackupFilesGetParameters()
'Overview                    : 当スクリプトの引数取得
'Detailed Description        : パラメータ格納用汎用ハッシュマップにKey="Parameter"で格納する
'                              個別パラメータ格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Seq(1,2,...)             個別パラメータ格納用ハッシュマップ
'                              引数があり存在するファイルパスのみ取得する
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFilesGetParameters( _
    byRef aoParams _
    )
    'パラメータ格納用ハッシュマップ
    Dim oParameter : Set oParameter = new_Dic()
    Dim lCnt : lCnt = 0
    Dim lFileFolderKbn : Dim sParam
    For Each sParam In WScript.Arguments
        'ファイルが存在する場合1、フォルダが存在する場合2
        lFileFolderKbn = 0
        If new_Fso().FileExists(sParam) Then lFileFolderKbn = 1
        If new_Fso().FolderExists(sParam) Then lFileFolderKbn = 2
        
        If lFileFolderKbn Then
        'ファイルまたはフォルダが存在する場合パラメータを取得
            lCnt = lCnt + 1
            Call oParameter.Add(lCnt, func_BackupFilesGetMapParameterInfo(lFileFolderKbn, sParam))
        End If
    Next
    
    Call aoParams.Add("Parameter", oParameter)
    
    'オブジェクトを開放
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_BackupFilesGetMapParameterInfo()
'Overview                    : 個別パラメータ格納用ハッシュマップ作成
'Detailed Description        : 個別パラメータ格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              "isFile"                 True:対象がファイル / False:対象がフォルダ
'                              "Path"                   フルパス
'Argument
'     alFileFolderKbn        : ファイルの場合1、フォルダの場合2
'     asPath                 : フルパス
'Return Value
'     個別パラメータ格納用ハッシュマップ
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
'Overview                    : バックアップする
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_BackupFilesBackup( _
    byRef aoParams _
    )
    'パラメータ格納用汎用ハッシュマップ
    Dim oParameter : Set oParameter = aoParams.Item("Parameter")
    
    Dim lKey
    For lKey=1 To oParameter.Count
    'パラメータごとのバックアップ処理
        Call sub_BackupFileBackupDetail(aoParams, oParameter.Item(lKey))
    Next
    
    'オブジェクトを開放
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : sub_BackupFileBackupDetail()
'Overview                    : パラメータごとのバックアップ処理
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     aoParameter            : 個別パラメータ格納用ハッシュマップ
'Return Value
'     なし
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
    'ファイルの場合
        Call sub_BackupFileProcForOneFile(aoParams, aoParameter.Item("Path"))
    Else
    'フォルダの場合
        Call sub_BackupFileProcForFolder(aoParams, aoParameter.Item("Path"))
    End If
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1
'Function/Sub Name           : sub_BackupFileProcForFolder()
'Overview                    : フォルダのバックアップ処理
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     asPath                 : バックアップ対象フォルダのフルパス
'Return Value
'     なし
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
    'フォルダ内のサブフォルダの処理
        Call sub_BackupFileProcForFolder(aoParams, oItem.Path)
    Next
    For Each oItem In func_CM_FsGetFiles(asPath)
    'フォルダ内のファイルの処理
        Call sub_BackupFileProcForOneFile(aoParams, oItem.Path)
    Next
    
    Set oItem = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-2
'Function/Sub Name           : sub_BackupFileProcForOneFile()
'Overview                    : ファイルごとのバックアップ処理
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     asPath                 : バックアップ対象ファイルのフルパス
'Return Value
'     なし
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
    
    '前バージョンのファイルを探して情報を取得する
    Call sub_BackupFileFindPreviousFile(aoParams, asPath)
    
    With aoParams.Item(asPath).Item("LatestHistoryInfo")
        '最新履歴がない または 最終更新日時が不一致 の場合はバックアップする
        Dim boDoBackup : boDoBackup = False
        If Not(.Item("Exists")) Then
            boDoBackup = True
        ElseIf .Item("DateLastModified") <> (func_CM_FsGetFile(asPath)).DateLastModified Then
            boDoBackup = True
        End If
        
        If Not(boDoBackup) Then
        'バックアップしない場合は関数を抜ける
            Exit Sub
        End If
        
        'バックアップファイル名の作成
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
    
    'コピー実施
    Dim sNewFilePath : sNewFilePath = func_CM_FsBuildPath(aoParams.Item(asPath).Item("OutputFolderPath"), sNewFileName)
    Call func_CM_FsCopyFile(asPath, sNewFilePath)
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-2-1
'Function/Sub Name           : sub_BackupFileFindPreviousFile()
'Overview                    : 前バージョンのファイルを探して情報を取得する
'Detailed Description        : パラメータ格納用汎用ハッシュマップに下記を格納する
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              asPathの値               バックアップ処理情報格納用ハッシュマップ
'                              
'                              バックアップ処理情報格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              OutputFolderPath         バックアップ出力先
'                              LatestHistoryInfo        最新バックアップ履歴格納用ハッシュマップ
'                              
'                              最新バックアップ履歴格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Exists                   True:最新履歴がある / False:最新履歴がない
'                              DateLastModified         最終更新日時
'                              Size                     サイズ
'                              BackupDate               履歴取得日（YYYYMMDD形式）
'                              Sequence                 履歴取得日が同日の場合の連番（1,2,3,...）
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     asPath                 : バックアップ対象ファイルのフルパス
'Return Value
'     なし
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
    
    'バックアップ先フォルダを探す
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
    
    'バックアップ対象ファイルのファイル名と拡張子からバックアップ履歴ファイル抽出用の正規表現を作成
    Dim sBasename : sBasename = func_CM_FsGetGetBaseName(asPath)
    Dim sExtensionName : sExtensionName = func_CM_FsGetGetExtensionName(asPath)
    Dim sPattern
    sPattern = sBasename & "_" & "(20)?(\d{2}[01]\d[0123]\d)" & "((_)(\d+))?"
    If (Len(sExtensionName)) Then sPattern = sPattern & "." & sExtensionName
    
    'バックアップ先フォルダから直近のファイル/フォルダを探す
    With New RegExp
        '初期化
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
            'バックアップ履歴の場合
                '名前から日付、連番部分を取得
                sDateToComp = .Replace(sItemName, "$2")
                If Len(sDateToComp)=6 Then sDateToComp = "20" & sDateToComp
                sSeqToComp = .Replace(sItemName, "$5")
                If (Len(sSeqToComp)=0) Then sSeqToComp = "1"
                
                If (sDateToComp > sDate) _
                    Or ((sDateToComp = sDate) And ( Clng(sSeqToComp) > lSeq )) _
                    Or (Len(sTargetPath)=0) Then
                '保持している情報より新しい場合、最新のバックアップ履歴として情報を取得
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
    
    'パラメータ格納用汎用ハッシュマップに格納する
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
