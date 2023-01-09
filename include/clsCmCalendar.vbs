'***************************************************************************************************
'FILENAME                    : clsCmCalendar.vbs
'Overview                    : 日付クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmCalendar()
'Overview                    : インスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     日付クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmCalendar( _
    )
    Set new_clsCmCalendar = (New clsCmCalendar).GetNow()
End Function

Class clsCmCalendar
    'クラス内変数、定数
    Private PdtNow
    Private PdbTimer
    
    'コンストラクタ
    Private Sub Class_Initialize()
        PdtNow = 0
        PdbTimer = 0
    End Sub
    'デストラクタ
    Private Sub Class_Terminate()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : GetNow()
    'Overview                    : 今の日付時刻を取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function GetNow( _
        )
        Set GetNow = func_CmCalendarGetNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : SetDateTime()
    'Overview                    : 指定した日付時刻を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avNow                  : 設定する日付時刻
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetDateTime( _
        ByVal avNow _
        )
        Set SetDateTime = func_CmCalendarSetDate(avNow)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : SetDateTimeWithFractionalSeconds()
    'Overview                    : 指定した日付時刻およびミリ秒を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avNow                  : 設定する日付時刻
    '     avTimer                : 設定するミリ秒（Timer()の値）
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function SetDateTimeWithFractionalSeconds( _
        ByVal avNow _
        , ByVal avTimer _
        )
        Set SetDateTimeWithFractionalSeconds = func_CmCalendarSetDateWithFractionalSeconds(avNow, avTimer)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : DisplayFormatAs()
    'Overview                    : 日付を整形する
    'Detailed Description        : func_CmCalendarSetDisplayFormatAs()に委譲する
    'Argument
    '     asFormat               : 表示形式
    'Return Value
    '     整形した日付
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function DisplayFormatAs( _
        ByVal asFormat _
        )
        DisplayFormatAs = func_CmCalendarSetDisplayFormatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : GetSerial()
    'Overview                    : シリアル値を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     シリアル値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function GetSerial( _
        )
       GetSerial = CDbl(Fix(PdtNow) + PdbTimer/(60*60*24))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : DifferenceInScondsFrom()
    'Overview                    : 差を秒数で返す
    'Detailed Description        : 工事中
    'Argument
    '     aoTarget               : 比較するclsCmCalendar型のインスタンス
    'Return Value
    '     差の秒数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function DifferenceInScondsFrom( _
        byRef aoTarget _
        )
        Dim dbDifference : dbDifference = CDbl((Me.GetSerial()-aoTarget.GetSerial())*60*60*24)
        DifferenceInScondsFrom = Fix(dbDifference) & "." _
                                 & func_CM_FillInTheCharacters( _
                                                              Fix( (dbDifference - Fix(dbDifference))*10^6 ) _
                                                              , 6 _
                                                              , "0" _
                                                              , False _
                                                              , True _
                                                              )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : CompareTo()
    'Overview                    : 日付の大小比較する
    'Detailed Description        : 下記比較結果を返す
    '                               0  引数と同値
    '                               -1 引数より小さい
    '                               1  引数より大きい
    'Argument
    '     aoTarget               : 比較するclsCmCalendar型のインスタンス
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/05         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function CompareTo( _
        byRef aoTarget _
        )
        Dim SerialMe : SerialMe = Me.GetSerial()
        Dim SerialTg : SerialTg = aoTarget.GetSerial()
        Dim lResult : lResult = 0
        If (SerialMe < SerialTg) Then lResult = -1
        If (SerialMe > SerialTg) Then lResult = 1
        CompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarGetNow()
    'Overview                    : 今の日付時刻を取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarGetNow( _
        )
        PdtNow = Now()
        PdbTimer = Timer()
        Set func_CmCalendarGetNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDate()
    'Overview                    : 指定した日付時刻を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avNow                  : 設定する日付時刻
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDate( _
        ByVal avNow _
        )
        On Error Resume Next
        PdtNow = CDate(avNow)
        PdbTimer = 0
        If Err.Number Then
            PdtNow = 0
            Err.Clear
        End If
        Set func_CmCalendarSetDate = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDateWithFractionalSeconds()
    'Overview                    : 指定した日付時刻とミリ秒を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avNow                  : 設定する日付時刻
    '     avTimer                : 設定するミリ秒（Timer()の値）
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDateWithFractionalSeconds( _
        ByVal avNow _
        , ByVal avTimer _
        )
        On Error Resume Next
        PdtNow = CDate(avNow)
        PdbTimer = CDbl(avTimer)
        If Err.Number Then
            PdtNow = 0
            PdbTimer = 0
            Err.Clear
        End If
        Set func_CmCalendarSetDateWithFractionalSeconds = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDisplayFormatAs()
    'Overview                    : 日付を整形する
    'Detailed Description        : 下記設定値は日付の数値が入る、下記以外の値はそのまま使用する
    '                              なお、日付が8の場合に"DD"は"08"、"D"は"8"を表示する
    '                              例） "YY/M/DD hh:mm:ss.000" → 23/1/04 16:55:12.345
    '                               YY[YY]    西暦年
    '                               M{M]      月
    '                               D{D]      日
    '                               h{h]      時
    '                               m{m]      分
    '                               s{s]      秒
    '                               .000      ミリ秒まで
    '                               .000000   ナノ秒まで
    'Argument
    '     asFormat               : 表示形式
    'Return Value
    '     整形した日付
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCalendarSetDisplayFormatAs( _
        byVal asFormat _
        )
        Dim oConversionSettings : Set oConversionSettings = CreateObject("Scripting.Dictionary")
        With oConversionSettings
            '変換テーブル定義
            .Add "YYYY", Array("UseDatePart()", "yyyy", False)
            .Add "yyyy", Array("UseDatePart()", "yyyy", False)
            .Add "YY", Array("UseDatePart()", "yyyy", True)
            .Add "yy", Array("UseDatePart()", "yyyy", True)
            .Add "MM", Array("UseDatePart()", "m", False)
            .Add "M", Array("UseDatePart()", "m", False)
            .Add "DD", Array("UseDatePart()", "d", False)
            .Add "dd", Array("UseDatePart()", "d", False)
            .Add "D", Array("UseDatePart()", "d", False)
            .Add "d", Array("UseDatePart()", "d", False)
            .Add "HH", Array("UseDatePart()", "h", False)
            .Add "hh", Array("UseDatePart()", "h", False)
            .Add "H", Array("UseDatePart()", "h", False)
            .Add "h", Array("UseDatePart()", "h", False)
            .Add "mm", Array("UseDatePart()", "n", False)
            .Add "m", Array("UseDatePart()", "n", False)
            .Add "SS", Array("UseDatePart()", "s", False)
            .Add "ss", Array("UseDatePart()", "s", False)
            .Add "S", Array("UseDatePart()", "s", False)
            .Add "s", Array("UseDatePart()", "s", False)
            .Add ".000000", Array("GetFractionalSeconds")
            .Add ".000", Array("GetFractionalSeconds")
            
            Dim lPos : lPos=1
            Dim sResult : sResult=""
            Dim lKeyLen : Dim boIsMatch : Dim sItemValue : Dim sKey : Dim vItem
            Do Until(Len(asFormat)<lPos)
                '初期化
                boIsMatch = False : sItemValue = ""
                
                'すべての変換テーブルの情報を確認する
                For Each sKey In .Keys
                    'キーの文字数を取得
                    lKeyLen=Len(sKey)
                    
                    If StrComp(sKey, Mid(asFormat, lPos, lKeyLen))=0 Then
                    '変換テーブルにある文字と一致した場合
                        vItem = .Item(sKey)
                        If vItem(0)="UseDatePart()" Then
                        'PdtNowからDatePart()で値を取り出す場合
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtNow), lKeyLen, "0", vItem(2), True)
                        Else
                        'PdbTimerからミリ秒をを取り出す場合
                            sItemValue = "." & func_CM_FillInTheCharacters(Fix((PdbTimer-Fix(PdbTimer))*10^(lKeyLen-1)), lKeyLen-1, "0", False, True)
                        End If
                        boIsMatch = True : Exit For
                    End If
                Next
                
                If boIsMatch Then
                '変換テーブルありの場合、マッチしたキーの文字数だけ進める
                    lPos=lPos+lKeyLen
                Else
                '変換テーブルなしの場合、asFormatの1文字をそのまま使用し1文字進める
                    sItemValue = Mid(asFormat, lPos, 1)
                    lPos=lPos+1
                End If
                sResult = sResult & sItemValue
            Loop
            
        End With
        func_CmCalendarSetDisplayFormatAs = sResult
        Set oConversionSettings = Nothing
    End Function
    
End Class
