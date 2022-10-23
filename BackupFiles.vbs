'***************************************************************************************************
'FILENAME                    : BackupFiles.vbs
'Overview                    : 引数で受け取ったファイルをバックアップする
'Detailed Description        : Sendtoから使用する
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
Private Const Cs_FOLDER_INCLUDE = "include"

'Include用関数定義
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
    
    Dim oParams : Set oParams = CreateObject("Scripting.Dictionary")
    
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
    Dim oParameter : Set oParameter = CreateObject("Scripting.Dictionary")
    Dim lCnt : lCnt = 0
    Dim lFileFolderKbn : Dim sParam
    For Each sParam In WScript.Arguments
        'ファイルが存在する場合1、フォルダが存在する場合2
        lFileFolderKbn = 0
        If func_CM_FsFileExists(sParam) Then lFileFolderKbn = 1
        If func_CM_FsFolderExists(sParam) Then lFileFolderKbn = 2
        
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
    'ファイルごとに処理する
        Call sub_BackupFileBackupDetail(aoParams, oParameter.Item(lKey))
    Next
    
    'オブジェクトを開放
    Set oParameter = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : sub_BackupFileBackupDetail()
'Overview                    : ファイルごとのバックアップ処理
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
    
    '前バージョンのファイルを探して情報を取得する
    Call sub_BackupFileFindPreviousFile(aoParams, aoParameter)
    
    Call Msgbox("OutputFolderPath : " & aoParams.Item("OutputFolderPath"))
    Call Msgbox("Path : " & aoParams.Item("LatestHistoryInfo").Item("Path"))
    Call Msgbox("Date : " & aoParams.Item("LatestHistoryInfo").Item("Date"))
    Call Msgbox("Sequence : " & aoParams.Item("LatestHistoryInfo").Item("Sequence"))
    'バックアップ要否判断
    
    'バックアップ実施
     'バックアップファイル名の確定
     'コピー実施
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1
'Function/Sub Name           : sub_BackupFileFindPreviousFile()
'Overview                    : 前バージョンのファイルを探して情報を取得する
'Detailed Description        : パラメータ格納用汎用ハッシュマップに下記を格納する
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              OutputFolderPath         バックアップ出力先のパス
'                              
'                              パラメータ格納用汎用ハッシュマップにKey="LatestHistoryInfo"で格納する
'                              最新バックアップ履歴格納用ハッシュマップの構成
'                              Key                      Value
'                              -------------------      --------------------------------------------
'                              Path                     パス
'                              Date                     履歴取得日（YYYYMMDD形式）
'                              Sequence                 連番（1,2,3,...）
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
Private Sub sub_BackupFileFindPreviousFile( _
    byRef aoParams _
    , byRef aoParameter _
    )
    
    'バックアップ先フォルダを探す
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
    
    'バックアップ対象ファイルのファイル名と拡張子からバックアップ履歴ファイル抽出用の正規表現を作成
    Dim sBasename : sBasename = func_CM_FsGetGetBaseName(aoParameter.Item("Path"))
    Dim sExtensionName : sExtensionName = func_CM_FsGetGetExtensionName(aoParameter.Item("Path"))
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
        For Each oItem In func_BackupFilesGetTargetList(aoParameter, sTargetFolder)
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
                    Or Not(Len(sTargetPath)) Then
                '保持している情報より新しい場合、最新のバックアップ履歴として情報を取得
                    sTargetPath = oItem.Path
                    sDate = sDateToComp
                    lSeq = Clng(sSeqToComp)
                End If
            End If
        Next
    End With
    
    'パラメータ格納用汎用ハッシュマップに格納する
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
'Overview                    : バックアップ先フォルダのファイルまたはフォルダのリストを取得
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     aoParameter            : 個別パラメータ格納用ハッシュマップ
'Return Value
'     FilesまたはFoldersコレクション
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
    
    'バックアップファイルから直近のファイル/フォルダを探す
    If aoParameter.Item("isFile") Then
        Set func_BackupFilesGetTargetList = func_CM_FsGetFiles(asTargetFolder)
    Else
        Set func_BackupFilesGetTargetList = func_CM_FsGetFolders(asTargetFolder)
    End If
    
End Function
