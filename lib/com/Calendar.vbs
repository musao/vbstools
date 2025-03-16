'***************************************************************************************************
'FILENAME                    : Calendar.vbs
'Overview                    : 日付クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Class Calendar
    'クラス内変数、定数
    Private PdtDateTime, PdbElapsedSeconds, PsDefaultFormat
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : コンストラクタ
    'Detailed Description        : 内部変数の初期化
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PdtDateTime = Null
        PdbElapsedSeconds = Null
        PsDefaultFormat = "YYYY/MM/DD hh:mm:ss.000"
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Terminate()
    'Overview                    : デストラクタ
    'Detailed Description        : 終了処理
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get dateTime()
    'Overview                    : 日時を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     日時
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get dateTime()
       dateTime = PdtDateTime
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get fractionalPartOfElapsedSeconds()
    'Overview                    : 経過秒の小数部を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     経過秒の小数部
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get fractionalPartOfElapsedSeconds()
       fractionalPartOfElapsedSeconds = Null
       If Not IsNull(PdtDateTime) Then fractionalPartOfElapsedSeconds = this_getfractionalPartOfElapsedSeconds()
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get elapsedSeconds()
    'Overview                    : 経過秒を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     経過秒
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get elapsedSeconds()
       elapsedSeconds = PdbElapsedSeconds
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get serial()
    'Overview                    : 日付のシリアル値を返す
    'Detailed Description        : シリアル値とは1900/1/1を1として、何日経過したかを示す数値
    'Argument
    '     なし
    'Return Value
    '     日付のシリアル値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get serial()
       serial = Null
       If Not IsNull(PdtDateTime) Then serial = Cdbl(PdtDateTime)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : デフォルトの形式で表示する
    'Detailed Description        : this_formatAs()に委譲する
    'Argument
    '     なし
    'Return Value
    '     デフォルトの形式に整形した日付
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get toString()
        toString = this_formatAs(PsDefaultFormat)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : clone()
    'Overview                    : 自身と同じ内容の新しいインスタンスを作る
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function clone( _
        )
        Set clone = this_clone()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
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
    Public Function compareTo( _
        byRef aoTarget _
        )
        ast_argsAreSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+compareTo()", "That object is not a calendar class."
        compareTo = this_compareTo(aoTarget)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : differenceFrom()
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
    Public Function differenceFrom( _
        byRef aoTarget _
        )
        ast_argsAreSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+differenceFrom()", "That object is not a calendar class."
        differenceFrom = this_differenceFrom(aoTarget)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : formatAs()
    'Overview                    : 日付を整形する
    'Detailed Description        : this_formatAs()に委譲する
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
    Public Function formatAs( _
        byVal asFormat _
        )
        formatAs = this_formatAs(asFormat)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : of()
    'Overview                    : 引数に応じたインスタンスを作成する
    'Detailed Description        : this_of()に委譲する
    'Argument
    '     avArgument             : 引数
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function of( _
        byRef avArgument _
        )
        Set of = this_of(avArgument, TypeName(Me)&"+of()")
    End Function
     
    '***************************************************************************************************
    'Function/Sub Name           : ofNow()
    'Overview                    : 今の日付時刻を取得する
    'Detailed Description        : this_setData()に委譲する
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
    Public Function ofNow( _
        )
        Set ofNow = this_setData(Now(), Timer(), TypeName(Me)&"+ofNow()")
    End Function
       



    
    '***************************************************************************************************
    'Function/Sub Name           : this_clone()
    'Overview                    : 自身と同じ内容の新しいインスタンスを作る
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_clone( _
        )
        Dim oNewIns : Set oNewIns = new Calendar
        If IsNull(PdtDateTime) Then
        Else
            If IsNull(PdbElapsedSeconds) Then
                Call oNewIns.of(Array(PdtDateTime))
            Else
                Call oNewIns.of(Array(PdtDateTime, PdbElapsedSeconds))
            End If
        End If
        Set this_clone = oNewIns
        Set oNewIns = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_compareTo()
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
    '2025/02/01         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_compareTo( _
        byRef aoTarget _
        )
        this_compareTo = 0
        If IsNull(PdtDateTime) And IsNull(aoTarget.dateTime) Then Exit Function
        
        Dim lResult : lResult = 0
        If IsNull(PdtDateTime) Or (PdtDateTime < aoTarget.dateTime) Then lResult = -1
        If IsNull(aoTarget.dateTime) Or (PdtDateTime > aoTarget.dateTime) Then lResult = 1
        If lResult <> 0 Then
            this_compareTo = lResult
            Exit Function
        End If
        
        If (this_getfractionalPartOfElapsedSeconds < aoTarget.fractionalPartOfElapsedSeconds) Then lResult = -1
        If (this_getfractionalPartOfElapsedSeconds > aoTarget.fractionalPartOfElapsedSeconds) Then lResult = 1
        this_compareTo = lResult

    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_differenceFrom()
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
    '2025/02/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function this_differenceFrom( _
        byRef aoTarget _
        )
        If this_compareTo(aoTarget)=0 Then
            this_differenceFrom = 0
            Exit function
        End If

        Dim dbResult : dbResult = 0
        If IsNull(PdtDateTime) Then dbResult = -1 * ((aoTarget.dateTime)*60*60*24 + aoTarget.fractionalPartOfElapsedSeconds)
        If IsNull(aoTarget.dateTime) Then dbResult = PdtDateTime*60*60*24 + this_getfractionalPartOfElapsedSeconds
        If dbResult <> 0 Then
            this_differenceFrom = dbResult
            Exit Function
        End If

        Dim dbDiffElapsedSeconds
        dbDiffElapsedSeconds = this_getfractionalPartOfElapsedSeconds-aoTarget.fractionalPartOfElapsedSeconds

        If (PdtDateTime <> aoTarget.dateTime) Then dbDiffElapsedSeconds = dbDiffElapsedSeconds+(PdtDateTime-aoTarget.dateTime)*60*60*24
        this_differenceFrom = math_round(dbDiffElapsedSeconds, 6)

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_formatAs()
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
    '                               000       ミリ秒まで
    '                               000000    マイクロ秒まで
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
    Private Function this_formatAs( _
        byVal asFormat _
        )
        this_formatAs = "<"&TypeName(Me)&">"&cf_toString(Null)
        If IsNull(PdtDateTime) Then Exit Function

        Const Cl_USE_DATAPART = 0
        Const Cl_USE_FRACTIONAL_SECONDS = 1
        With new_Dic()
            '変換テーブル定義
            .Add "YYYY", Array(Cl_USE_DATAPART, "yyyy", False)
            .Add "yyyy", Array(Cl_USE_DATAPART, "yyyy", False)
            .Add "YY", Array(Cl_USE_DATAPART, "yyyy", True)
            .Add "yy", Array(Cl_USE_DATAPART, "yyyy", True)
            .Add "MM", Array(Cl_USE_DATAPART, "m", False)
            .Add "M", Array(Cl_USE_DATAPART, "m", False)
            .Add "DD", Array(Cl_USE_DATAPART, "d", False)
            .Add "dd", Array(Cl_USE_DATAPART, "d", False)
            .Add "D", Array(Cl_USE_DATAPART, "d", False)
            .Add "d", Array(Cl_USE_DATAPART, "d", False)
            .Add "HH", Array(Cl_USE_DATAPART, "h", False)
            .Add "hh", Array(Cl_USE_DATAPART, "h", False)
            .Add "H", Array(Cl_USE_DATAPART, "h", False)
            .Add "h", Array(Cl_USE_DATAPART, "h", False)
            .Add "mm", Array(Cl_USE_DATAPART, "n", False)
            .Add "m", Array(Cl_USE_DATAPART, "n", False)
            .Add "SS", Array(Cl_USE_DATAPART, "s", False)
            .Add "ss", Array(Cl_USE_DATAPART, "s", False)
            .Add "S", Array(Cl_USE_DATAPART, "s", False)
            .Add "s", Array(Cl_USE_DATAPART, "s", False)
            .Add "000000", Array(Cl_USE_FRACTIONAL_SECONDS)
            .Add "000", Array(Cl_USE_FRACTIONAL_SECONDS)
            
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
                    
                    If cf_isSame(sKey, Mid(asFormat, lPos, lKeyLen)) Then
                    '変換テーブルにある文字と一致した場合
                        vItem = .Item(sKey)
                        If cf_isSame(Cl_USE_DATAPART, vItem(0)) Then
                        'PdtDateからDatePart()で値を取り出す場合
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDateTime), lKeyLen, "0", vItem(2), True)
                        Else
                        '秒数の小数部を取り出す場合
                            sItemValue = func_CM_FillInTheCharacters(math_tranc(this_getfractionalPartOfElapsedSeconds*10^lKeyLen), lKeyLen, "0", False, True)
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
        this_formatAs = sResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getfractionalPartOfElapsedSeconds()
    'Overview                    : 経過秒の小数部を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     経過秒の小数部
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getfractionalPartOfElapsedSeconds( _
        )
        Dim dbFract : dbFract = 0
        If Not IsNull(PdbElapsedSeconds) Then dbFract = math_round(math_fractional(PdbElapsedSeconds),7)
        this_getfractionalPartOfElapsedSeconds = dbFract
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_of()
    'Overview                    : 引数に応じたインスタンスを作成する
    'Detailed Description        : this_setData()に委譲する
    '                              以下の入力検査を行う
    '                              1.配列でない場合
    '                                Date型（小数点以下の秒数があってもよい）
    '                              2.配列の場合は要素数に応じたチェックを行う
    '                                1-1.要素数が1つ
    '                                 e(0) -> Date型（小数点以下の秒数があってもよい）
    '                                1-2.要素数が2つ
    '                                 e(0) -> Date型
    '                                 e(1) -> Double型
    '                                1-3.要素数が6つ
    '                                 e(0-5) -> "e(0)/e(1)/e(2) e(3):e(4):e(5)"がDate型
    '                                1-4.要素数が7つ
    '                                 e(0-5) -> "e(0)/e(1)/e(2) e(3):e(4):e(5)"がDate型
    '                                 e(6) -> Double型
    '                                1-5.上記以外の要素数はエラーとする
    'Argument
    '     avArgument             : 引数
    '     asSource               : ソース
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_of( _
        byRef avArgument _
        , byVal asSource _
        )
        Dim dtDateTime, dbElapsedSeconds, boIsError
        dtDateTime = Null
        dbElapsedSeconds = Null
        boIsError = False
        
        On Error Resume Next
        If Not(IsArray(avArgument)) Then
        '配列でない場合
            Call this_ofForOneArg(avArgument, dtDateTime, dbElapsedSeconds)
        ElseIf new_Arr().hasElement(avArgument) Then
        '配列の要素がある場合
            Dim e : e = avArgument
            Select Case Ubound(e)
                Case 0:
                    Call this_ofForOneArg(e(0), dtDateTime, dbElapsedSeconds)
                Case 1:
                    dtDateTime = Cdate(e(0))
                    dbElapsedSeconds = Cdbl(e(1))
                Case 5:
                    dtDateTime = Cdate(e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5))
                Case 6:
                    dtDateTime = Cdate(e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5))
                    dbElapsedSeconds = Cdbl(e(6))
            End Select
        End If
        If Err.Number<>0 Then boIsError=True
        On Error Goto 0

        ast_argFalse boIsError, asSource, "invalid argument. " & cf_toString(avArgument)

        Set this_of = this_setData(dtDateTime, dbElapsedSeconds, asSource)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_ofForOneArg()
    'Overview                    : 日付型に変換する
    'Detailed Description        : 工事中
    'Argument
    '     avDateTime             : 引数の日付時刻
    '     adtDateTime            : 日時
    '     dbElapsedSeconds       : 経過秒
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_ofForOneArg( _
        byRef avDateTime _
        , byRef adtDateTime _
        , byRef adbElapsedSeconds _
        )
        Dim oRe : Set oRe = new_Re("^([^.]+)\.(\d+)$", "")
        If oRe.Test(avDateTime) Then
            adtDateTime = Cdate(oRe.Replace(avDateTime, "$1"))
            Dim dbElapsedSecondsByDt : dbElapsedSecondsByDt = math_tranc(math_fractional(adtDateTime)*24*60*60)
            adbElapsedSeconds = dbElapsedSecondsByDt + Cdbl("0." & oRe.Replace(avDateTime, "$2"))
        Else
            adtDateTime = Cdate(avDateTime)
        End If
        Set oRe = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setData()
    'Overview                    : データを設定する
    'Detailed Description        : 工事中
    'Argument
    '     adtDateTime            : 日時
    '     adbElapsedSeconds      : 経過秒
    '     asSource               : ソース
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setData( _
        byVal adtDateTime _
        , byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argNull PdtDateTime, asSource, "Because it is an immutable variable, its value cannot be changed."
        this_setDateTime adtDateTime, asSource
        this_setElapsedSeconds adbElapsedSeconds, asSource
        Set this_setData = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_setDateTime()
    'Overview                    : PadtDateTimeのセッター
    'Detailed Description        : 工事中
    'Argument
    '     adtDateTime            : 日時
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setDateTime( _
        byVal adtDateTime _
        , byVal asSource _
        )
        ast_argTrue IsDate(adtDateTime), asSource, "DateTime is not a date/time."
        PdtDateTime = Cdate(adtDateTime)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setElapsedSeconds()
    'Overview                    : PadbElapsedSecondsのセッター
    'Detailed Description        : 工事中
    'Argument
    '     adbElapsedSeconds      : 経過秒
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setElapsedSeconds( _
        byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argTrue (IsNull(adbElapsedSeconds) Or cf_isNonNegativeNumber(adbElapsedSeconds)), asSource, "ElapsedSeconds must be null or a non-negative number."
        If Not(IsNull(adbElapsedSeconds)) Then PdbElapsedSeconds = Cdbl(adbElapsedSeconds)
    End Sub
    
End Class
