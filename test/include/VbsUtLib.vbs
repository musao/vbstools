'***************************************************************************************************
'FILENAME                    : VbsUrLib.vbs
'Overview                    : 単体テスト用ライブラリ
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : sub_UtResultOutput()
'Overview                    : UT結果を出力する
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/13         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_UtResultOutput(_
    byRef aoUtAssistant _
    )
    
    With aoUtAssistant
        'ログファイル出力
        Call sub_UtWriteFile(func_UtGetThisLogFilePath(), .OutputReportInTsvFormat())
        
        '結果をメッセージで出力
        Dim sMsg : sMsg = "NGがあります、ログを確認ください"
        If .isAllOk Then sMsg = "全ケースOKです！"
        Call Msgbox(sMsg)
    End With
    
End sub

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisWorkFolderPath()
'Overview                    : UT対象のソースファイル用のワークディレクトリのフルパスを取得
'Detailed Description        : ディレクトリがない場合は作成する
'Argument
'     なし
'Return Value
'     UT対象のソースファイル用のワークディレクトリのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisWorkFolderPath()
    With CreateObject("Scripting.FileSystemObject")
        Dim sThisWorkFolderPath
        sThisWorkFolderPath = .BuildPath( _
                                        .GetParentFolderName(WScript.ScriptFullName) _
                                        , .GetBaseName(WScript.ScriptFullName) _
                                        )
        If Not(.FolderExists(sThisWorkFolderPath)) Then .CreateFolder(sThisWorkFolderPath)
        func_UtGetThisWorkFolderPath = sThisWorkFolderPath
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisTempFilePath()
'Overview                    : 一時ファイルのフルパスを取得
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     一時ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisTempFilePath()
    With CreateObject("Scripting.FileSystemObject")
        func_UtGetThisTempFilePath = .BuildPath( _
                                        func_UtGetThisWorkFolderPath() _
                                        , .GetTempName() _
                                        )
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetThisLogFilePath()
'Overview                    : ログファイルのフルパスを取得
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     ログファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetThisLogFilePath()
    With CreateObject("Scripting.FileSystemObject")
        func_UtGetThisLogFilePath = .BuildPath( _
                                        func_UtGetThisWorkFolderPath() _
                                        , .GetBaseName(WScript.ScriptFullName) _
                                            & "_" & func_UtGetGetDateInYyyymmddhhmmssFormat(Now()) _
                                            & ".log" _
                                        )
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_UtGetGetDateInYyyymmddhhmmssFormat()
'Overview                    : 日時をYYYYMMDD_HHMMSS形式で取得する
'Detailed Description        : 工事中
'Argument
'     adtDate                : 日時
'Return Value
'     YYYYMMDD_HHMMSS形式の文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_UtGetGetDateInYyyymmddhhmmssFormat(_
    byVal adtDate _
    )
    Dim dtNow : dtNow = adtDate
    Dim sCont : sCont = Year(dtNow)
    sCont = sCont & Right("0" & Month(dtNow) , 2)
    sCont = sCont & Right("0" & Day(dtNow) , 2)
    sCont = sCont & "_"
    sCont = sCont & Right("0" & Hour(dtNow) , 2)
    sCont = sCont & Right("0" & Minute(dtNow) , 2)
    sCont = sCont & Right("0" & Second(dtNow) , 2)
    func_UtGetGetDateInYyyymmddhhmmssFormat = sCont
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_UtWriteFile()
'Overview                    : ファイル出力する
'Detailed Description        : エラーは無視する
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_UtWriteFile(_
    byVal asPath _
    , byVal asCont _
    )
    On Error Resume Next
    With CreateObject("Scripting.FileSystemObject")
        Call .OpenTextFile(asPath, 8, True).WriteLine(asCont)
    End With
    If Err.Number Then
        Err.Clear
    End If
End sub
