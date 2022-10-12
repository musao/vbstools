'***************************************************************************************************
'FILENAME                    : clsUtAssistant.vbs
'Overview                    : 単体テスト用アシスタントクラス
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
Class clsUtAssistant
'    'クラス内変数
    Private PoRecord
    Private PoRecordTitles
    
    'コンストラクタ
    Private Sub Class_Initialize()
        '結果格納用ハッシュマップ
        Set PoRecord = CreateObject("Scripting.Dictionary")
        '結果詳細ハッシュマップに格納する情報のタイトル定義
        Set PoRecordTitles = CreateObject("Scripting.Dictionary")
        Call PoRecordTitles.Add(1, "Seq")
        Call PoRecordTitles.Add(2, "CaseName")
        Call PoRecordTitles.Add(3, "Result")
        Call PoRecordTitles.Add(4, "Start")
        Call PoRecordTitles.Add(5, "End")
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoRecord = Nothing
    End Sub
    
    Public Property Get Record()
        Set Record = PoRecord
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Run()
    'Overview                    : テスト実施
    'Detailed Description        : 結果格納用ハッシュマップの構成
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              Seq(1,2,3…)              結果詳細ハッシュマップ
    '
    '                              結果詳細ハッシュマップの構成
    '                              Key                      Value
    '                              -------------------      --------------------------------------------
    '                              "CaseName"                実行するケース名（関数名）
    '                              "Result"                  結果 True,Flase
    '                              "Start"                   開始時刻
    '                              "End"                     終了時刻
    'Argument
    '     asCaseName             : 実行するケース名（関数名）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Run( _
        byVal asCaseName _
        )
        '実施
        Dim dtStart : dtStart = func_GetDateInMilliseconds()
        On Error Resume Next
        Dim boResult : boResult = GetRef(asCaseName)
        Dim dtEnd : dtEnd = func_GetDateInMilliseconds()
        If Err.Number Or Not(boResult) Then boResult = False
        
        '結果を記録
        Dim lSeq : lSeq = PoRecord.Count+1
        Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
        With PoRecordTitles
            Call oTemp.Add(.Item(1), lSeq)
            Call oTemp.Add(.Item(2), asCaseName)
            Call oTemp.Add(.Item(3), boResult)
            Call oTemp.Add(.Item(4), dtStart)
            Call oTemp.Add(.Item(5), dtEnd)
        End With
        Call PoRecord.Add(lSeq, oTemp)
        
        Set oTemp = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : OutputReportInTsvFormat()
    'Overview                    : 結果をTsv形式で出力する
    'Detailed Description        : 工事中
    'Argument
    '     asCaseName             : 実行するケース名（関数名）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function OutputReportInTsvFormat( _
        )
        
        Dim sDelimiter : sDelimiter = vbTab
        Dim sLineFeedCode : sLineFeedCode = vbCrLf
        Dim lKeyT : Dim lKeyC
        
        'ヘッダの編集
        Dim sHeader : sHeader = ""
        For lKeyT=1 To PoRecordTitles.Count
        'ヘッダはタイトル定義の内容を順に出力する
            If Len(sHeader) Then sHeader = sHeader & sDelimiter
            sHeader = sHeader & PoRecordTitles.Item(lKeyT)
        Next
        
        '内容の編集
        Dim sContLine
        Dim sCont : sCont = ""
        For lKeyC=1 To PoRecord.Count
        '内容は結果格納用ハッシュマップを順に処理する
            sContLine = ""
            For lKeyT=1 To PoRecordTitles.Count
            '結果ごとにタイトルをキーに値を取り出す
                If Len(sContLine) Then sContLine = sContLine & sDelimiter
                sContLine = sContLine & PoRecord.Item(lKeyC).Item(PoRecordTitles.Item(lKeyT))
            Next
            If Len(sCont) Then sCont = sCont & sLineFeedCode
            sCont = sCont & sContLine
        Next
        
        '編集結果を返却
        OutputReportInTsvFormat = sHeader & sLineFeedCode & sCont
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : isAllOk()
    'Overview                    : 全UTが成功したかどうかを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True,Flase
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function isAllOk( _
        )
        isAllOk = False
        Dim lKey
        For lKey=1 To PoRecord.Count
        '結果格納用ハッシュマップを順に確認する、Falseがあれば終了する
            if Not(PoRecord.Item(lKey).Item(PoRecordTitles.Item(3))) Then Exit Function
        Next
        isAllOk = True
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_GetDateInMilliseconds()
    'Overview                    : 現在日時をミリ秒で取得する
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
    Private Function func_GetDateInMilliseconds()
        Dim dtNowTime        '現在時刻
        Dim lHour            '時
        Dim lngMinute        '分
        Dim lngSecond        '秒
        Dim lngMilliSecond   'ミリ秒

        dtNowTime = Timer
        lMilliSecond = dtNowTime - Fix(dtNowTime)
        lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
        dtNowTime = Fix(dtNowTime)
        lSecond = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lMinute = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lHour = Right("0" & dtNowTime, 2)

        func_GetDateInMilliseconds = Date() & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
    End Function
    
End Class
