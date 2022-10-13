'***************************************************************************************************
'FILENAME                    : clsUtAssistant.vbs
'Overview                    : 単体テスト用アシスタントクラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Class clsUtAssistant
'    'クラス内変数
    Private PdtNow
    Private PdtDate
    Private PdtStart
    Private PdtEnd
    Private PoRecDetail
    Private PoRecDetailTitles
    Private PoRecSumTitles
    
    'コンストラクタ
    Private Sub Class_Initialize()
        '開始日時の取得
        PdtNow = Now
        PdtDate = Date
        PdtStart = Timer
        '結果サマリーのタイトル定義
        Set PoRecSumTitles = CreateObject("Scripting.Dictionary")
        With PoRecSumTitles
            Call .Add(1, "Result")
            Call .Add(2, "CaseCount")
            Call .Add(3, "OkCaseCount")
            Call .Add(4, "NgCaseCount")
            Call .Add(5, "Start")
            Call .Add(6, "End")
            Call .Add(7, "ElapsedTime")
        End With
        '結果格納用ハッシュマップ
        Set PoRecDetail = CreateObject("Scripting.Dictionary")
        '結果詳細ハッシュマップに格納する情報のタイトル定義
        Set PoRecDetailTitles = CreateObject("Scripting.Dictionary")
        With PoRecDetailTitles
            Call .Add(1, "Seq")
            Call .Add(2, "CaseName")
            Call .Add(3, "Result")
            Call .Add(4, "Start")
            Call .Add(5, "End")
            Call .Add(6, "ElapsedTime")
        End With
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
        Set PoRecDetailTitles = Nothing
        Set PoRecDetail = Nothing
        Set PoRecSumTitles = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCount()
    'Overview                    : 実施した単体テストケース数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     実施した全ケース数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCount()
        CaseCount = PoRecDetail.Count
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCountOk()
    'Overview                    : 実施した単体テストケースのうち成功した数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     実施した単体テストケースのうち成功した数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCountOk()
        CaseCountOk = func_CountCaseAs(True)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get CaseCountNg()
    'Overview                    : 実施した単体テストケースのうち失敗した数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     実施した単体テストケースのうち失敗した数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get CaseCountNg()
        CaseCountNg = func_CountCaseAs(False)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get StartTime()
    'Overview                    : 単体テストの開始日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     単体テストの開始日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get StartTime()
        StartTime = func_GetDateInMilliseconds(PdtDate, PdtStart)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ProcDate()
    'Overview                    : 単体テストの実施日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     単体テストの実施日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ProcDate()
        ProcDate = PdtNow
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get EndTime()
    'Overview                    : 単体テストの終了日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     最後の単体テストの終了日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get EndTime()
        EndTime = func_GetDateInMilliseconds(PdtDate, PdtEnd)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get ElapsedTime()
    'Overview                    : 単体テスト実施にかかった時間を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     単体テスト実施にかかった時間
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get ElapsedTime()
       ElapsedTime = PdtEnd - PdtStart
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
        Dim dtDate : dtDate = Date()
        Dim dtStart : dtStart = Timer
        On Error Resume Next
        Dim boResult : boResult = GetRef(asCaseName)
        Dim dtEnd : dtEnd = Timer
        If Err.Number Or Not(boResult) Then boResult = False
        
        '結果を記録
        Dim lSeq : lSeq = PoRecDetail.Count+1
        Dim oTemp : Set oTemp = CreateObject("Scripting.Dictionary")
        With PoRecDetailTitles
            Call oTemp.Add(.Item(1), lSeq)
            Call oTemp.Add(.Item(2), asCaseName)
            Call oTemp.Add(.Item(3), boResult)
            Call oTemp.Add(.Item(4), func_GetDateInMilliseconds(dtDate, dtStart))
            Call oTemp.Add(.Item(5), func_GetDateInMilliseconds(dtDate, dtEnd))
            Call oTemp.Add(.Item(6), dtEnd-dtStart)
        End With
        Call PoRecDetail.Add(lSeq, oTemp)
        
        '終了時間の取得
        PdtEnd = dtEnd
        
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
        
        'サマリー部
        Dim sSum : sSum = ""
        With PoRecSumTitles
            sSum = sSum & .Item(1) & sDelimiter & isAllOk & sLineFeedCode
            sSum = sSum & .Item(2) & sDelimiter & CaseCount & sLineFeedCode
            sSum = sSum & .Item(3) & sDelimiter & CaseCountOk & sLineFeedCode
            sSum = sSum & .Item(4) & sDelimiter & CaseCountNg & sLineFeedCode
            sSum = sSum & .Item(5) & sDelimiter & StartTime & sLineFeedCode
            sSum = sSum & .Item(6) & sDelimiter & EndTime & sLineFeedCode
            sSum = sSum & .Item(7) & sDelimiter & ElapsedTime & sLineFeedCode
        End With
        
        
        '詳細部
        Dim lKeyT : Dim lKeyC
        
        'ヘッダの編集
        Dim sHeader : sHeader = ""
        For lKeyT=1 To PoRecDetailTitles.Count
        'ヘッダはタイトル定義の内容を順に出力する
            If Len(sHeader) Then sHeader = sHeader & sDelimiter
            sHeader = sHeader & PoRecDetailTitles.Item(lKeyT)
        Next
        
        '内容の編集
        Dim sContLine
        Dim sCont : sCont = ""
        For lKeyC=1 To PoRecDetail.Count
        '内容は結果格納用ハッシュマップを順に処理する
            sContLine = ""
            For lKeyT=1 To PoRecDetailTitles.Count
            '結果ごとにタイトルをキーに値を取り出す
                If Len(sContLine) Then sContLine = sContLine & sDelimiter
                sContLine = sContLine & PoRecDetail.Item(lKeyC).Item(PoRecDetailTitles.Item(lKeyT))
            Next
            If Len(sCont) Then sCont = sCont & sLineFeedCode
            sCont = sCont & sContLine
        Next
        
        '編集結果を返却
        OutputReportInTsvFormat = sSum & sLineFeedCode & sHeader & sLineFeedCode & sCont
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : isAllOk()
    'Overview                    : 全単体テストが成功したかどうかを返す
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
        isAllOk = (PoRecDetail.Count=func_CountCaseAs(True))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_GetDateInMilliseconds()
    'Overview                    : 日時をミリ秒で取得する
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
    Private Function func_GetDateInMilliseconds( _
        byVal adtDate _
        , byVal adtTimer _
        )
        Dim dtNowTime        '現在時刻
        Dim lHour            '時
        Dim lngMinute        '分
        Dim lngSecond        '秒
        Dim lngMilliSecond   'ミリ秒

        dtNowTime = adtTimer
        lMilliSecond = dtNowTime - Fix(dtNowTime)
        lMilliSecond = Right("000" & Fix(lMilliSecond * 1000), 3)
        dtNowTime = Fix(dtNowTime)
        lSecond = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lMinute = Right("0" & dtNowTime Mod 60, 2)
        dtNowTime = dtNowTime \ 60
        lHour = Right("0" & dtNowTime, 2)

        func_GetDateInMilliseconds = adtDate & " " & lHour & ":" & lMinute & ":" & lSecond & "." & lMilliSecond
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CountCaseAs()
    'Overview                    : 結果ごとにケース数を数える
    'Detailed Description        : 工事中
    'Argument
    '     aboResult              : 数える対象のケース結果 True,Flase
    'Return Value
    '     ケース数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2022/10/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CountCaseAs( _
        byVal aboResult _
        )
        Dim lKey : Dim lCnt : lCnt = 0
        For lKey=1 To PoRecDetail.Count
        '結果格納用ハッシュマップから対象のケースを数える
            if PoRecDetail.Item(lKey).Item(PoRecDetailTitles.Item(3)) = aboResult Then lCnt = lCnt + 1
        Next
        func_CountCaseAs = lCnt
    End Function
    
End Class
