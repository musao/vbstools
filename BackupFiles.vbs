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
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'lib\com import
Dim sRelativeFolderName : sRelativeFolderName = "lib\com"
With CreateObject("Scripting.FileSystemObject")
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim sLibFolderPath : sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    Dim oLibFile
    For Each oLibFile In CreateObject("Shell.Application").Namespace(sLibFolderPath).Items
        If Not oLibFile.IsFolder Then
            If StrComp(.GetExtensionName(oLibFile.Path), "vbs", vbTextCompare)=0 Then ExecuteGlobal .OpenTextfile(oLibFile.Path).ReadAll
        End If
    Next
End With
Set oLibFile = Nothing
'lib import
sRelativeFolderName = "lib"
With new_FSO()
    sLibFolderPath = .BuildPath(sParentFolderPath, sRelativeFolderName)
    ExecuteGlobal .OpenTextfile(.BuildPath(sLibFolderPath,"libEnum.vbs")).ReadAll
End With


'ログ出力先、ブローカークラスのインスタンスの設定
Private PoTs4Log, PoBroker
Set PoTs4Log = fw_getTextstreamForLog()
Set PoBroker = new_BrokerOf(Array(topic.LOG, GetRef("this_logger")))

'Main関数実行
Call Main()

'終了処理
PoTs4Log.close()
Set PoBroker = Nothing : Set PoTs4Log = Nothing
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
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    'パラメータ格納用オブジェクト宣言
    Dim oParams : Set oParams = new_Dic()
    
    '当スクリプトの引数をパラメータ格納用オブジェクトに取得する
    fw_excuteSub "this_getParameters", oParams, PoBroker
    
    'バックアップする
    fw_excuteSub "this_backup", oParams, PoBroker
    
    'オブジェクトを開放
    Set oParams = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : this_getParameters()
'Overview                    : 当スクリプトの引数をパラメータ格納用オブジェクトに取得する
'Detailed Description        : パラメータ格納用汎用オブジェクトにKey="Param"で格納する
'                              配列（ArrayList型）に名前なし引数（/Key:Value 形式でない）を全て取得する
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_getParameters( _
    byRef aoParams _
    )
    'オリジナルの引数を取得
    Dim oArg : Set oArg = fw_storeArguments()
    '★ログ出力
    this_logger Array(logType.TRACE, "this_getParameters()", cf_toString(oArg))
    
    '実在するパスだけパラメータ格納用オブジェクトに設定
    Dim oParam, oItem : Set oParam = new_Arr()
    For Each oItem In oArg.Item("Unnamed")
        '引数からファイルシステムプロキシオブジェクトを生成する
        With fw_try(Getref("new_FspOf"), oItem)
            If Not .isErr() Then
                oParam.push .returnValue
            Else
                '★ログ出力
                this_logger Array(logType.WARNING, "this_getParameters()", oItem & " is an invalid argument.")
            End If
        End With
    Next
    cf_bindAt aoParams, "Param", oParam
    
    Set oItem = Nothing
    Set oParam = Nothing
    Set oArg = Nothing
End Sub
'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : this_backup()
'Overview                    : バックアップする
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用オブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_backup( _
    byRef aoParams _
    )
    'パラメータ格納用オブジェクト
    Dim oParam : Set oParam = aoParams.Item("Param").slice(0,Null)
    
    '個別のパラメータごとにバックアップを行う
    Do While oParam.length>0
        this_backupProc oParam.pop()
    Loop
    
    'オブジェクトを開放
    Set oParam = Nothing
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : this_backupProc()
'Overview                    : パラメータごとのバックアップ処理
'Detailed Description        : 工事中
'Argument
'     aoTarget               : backup対象（FileSystemProxy型のインスタンス）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_backupProc( _
    byRef aoTarget _
    )
    '★ログ出力
    this_logger Array(logType.INFO, "this_backupProc()", "Start backing up '" & aoTarget.toString() & "'.")

    If aoTarget.isFolder() Then
    'フォルダの場合
'        this_backupProcForFolder aoTarget
    Else
    'ファイルの場合
 '       this_backupProcForFile aoTarget
    End If

    '★ログ出力
    this_logger Array(logType.INFO, "this_backupProc()", "Backup of '" & aoTarget.toString() & "' has finished.")
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1-1
'Function/Sub Name           : this_backupProcForFolder()
'Overview                    : フォルダのバックアップ処理
'Detailed Description        : 工事中
'Argument
'     aoParams               : パラメータ格納用汎用ハッシュマップ
'     asPath                 : バックアップ対象フォルダのフルパス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_backupProcForFolder( _
    byRef aoParams _
    , byVal asPath _
    )
    
    Dim oItem
    For Each oItem In new_Fso().GetFolder(asPath).SubFolders
'    For Each oItem In func_CM_FsGetFolders(asPath)
    'フォルダ内のサブフォルダの処理
        Call sub_BackupFileProcForFolder(aoParams, oItem.Path)
    Next
    For Each oItem In new_Fso().GetFolder(asPath).Files
'    For Each oItem In func_CM_FsGetFiles(asPath)
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
'History
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
        ElseIf .Item("DateLastModified") <> (new_FileOf(asPath)).DateLastModified Then
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
        sNewFileName = sNewFileName & "." & new_Fso().GetExtensionName(asPath)
        
    End With
    
    'コピー実施
    Dim sNewFilePath : sNewFilePath = new_Fso().BuildPath(aoParams.Item(asPath).Item("OutputFolderPath"), sNewFileName)
    Call fs_copyFile(asPath, sNewFilePath)
    
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
'History
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
        sTemp = new_Fso().BuildPath(sParentFolderPath, oFolders.Item(lKey))
        If new_Fso().FolderExists(sTemp) Then
            sTargetFolder = sTemp
            Exit For
        End If
    Next
    If (Len(sTargetFolder)=0) Then
        sTargetFolder = new_Fso().BuildPath(sParentFolderPath, oFolders.Item(1))
        fs_createFolder sTargetFolder
    End If
    
    'バックアップ対象ファイルのファイル名と拡張子からバックアップ履歴ファイル抽出用の正規表現を作成
    Dim sBasename : sBasename = func_CM_FsGetGetBaseName(asPath)
    Dim sExtensionName : sExtensionName = new_Fso().GetExtensionName(asPath)
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
        For Each oItem In new_Fso().GetFolder(sTargetFolder).Files
'        For Each oItem In func_CM_FsGetFiles(sTargetFolder)
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
    If (Len(sTargetPath)>0) Then Set oTargetFile = new_FileOf(sTargetPath)
    
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

'***************************************************************************************************
'Processing Order            : -
'Function/Sub Name           : this_logger()
'Overview                    : ログ出力する
'Detailed Description        : fw_logger()に委譲する
'Argument
'     avParams               : 配列型のパラメータリスト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2026/01/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub this_logger( _
    byRef avParams _
    )
    fw_logger avParams, PoTs4Log
End Sub
