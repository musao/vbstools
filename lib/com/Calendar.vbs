'***************************************************************************************************
'FILENAME                    : Calendar.vbs
'Overview                    : 日付クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Class Calendar
    'クラス内変数、定数
    Private PdtDateTime, PdbElapsedSeconds, PsDefaultFormat, Cl_NUMBER_OF_SECONDS_IN_A_DAY
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : コンストラクタ
    'Detailed Description        : 内部変数の初期化
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/08/20         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PdtDateTime = Null
        PdbElapsedSeconds = Null
        PsDefaultFormat = "YYYY/MM/DD hh:mm:ss.000"
        Cl_NUMBER_OF_SECONDS_IN_A_DAY = 24 * 60 * 60 '24時間分の秒数
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
    'History
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
    'History
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get fractionalPartOfElapsedSeconds()
       fractionalPartOfElapsedSeconds = Null
       If Not this_isInitial() Then fractionalPartOfElapsedSeconds = this_getfractionalPartOfElapsedSeconds()
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
    'History
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get serial()
       serial = Null
       If Not this_isInitial() Then serial = Cdbl(PdtDateTime)
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get toString()
        toString = this_formatAs(PsDefaultFormat)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : addMilliSeconds()
    'Overview                    : 指定ミリ秒数を加算または減算した新しいインスタンスを返す
    'Detailed Description        : this_addDate()に委譲する
    'Argument
    '     avVal                  : 指定ミリ秒数
    'Return Value
    '     指定ミリ秒数を加算または減算した新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addMilliSeconds( _
        byVal avVal _
        )
        cf_bind addMilliSeconds, this_addDate("ms", avVal, TypeName(Me)&"+addMilliSeconds()")
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : addSeconds()
    'Overview                    : 指定秒数を加算または減算した新しいインスタンスを返す
    'Detailed Description        : this_addDate()に委譲する
    'Argument
    '     avVal                  : 指定秒数
    'Return Value
    '     指定秒数を加算または減算した新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addSeconds( _
        byVal avVal _
        )
        cf_bind addSeconds, this_addDate("s", avVal, TypeName(Me)&"+addSeconds()")
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : clone()
    'Overview                    : 自身と同じ内容の新しいインスタンスを作る
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
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
    'History
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
    'History
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
    'History
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
    '     avArg                  : 引数
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function of( _
        byRef avArg _
        )
        Set of = this_of(avArg, TypeName(Me)&"+of()")
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ofNow()
    'Overview                    : 今の日付時刻を取得する
    'Detailed Description        : this_of()に委譲する
    'Argument
    '     なし
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function ofNow( _
        )
        Set ofNow = this_of(Array(Now(), Timer()), TypeName(Me)&"+ofNow()")
    End Function


    

    
    '***************************************************************************************************
    'Function/Sub Name           : this_addDate()
    'Overview                    : 指定された日付に、指定した時間間隔（年、月、日、時間など）を加算
    '                              または減算して、新しい日付のインスタンスを返す
    'Detailed Description        : 工事中
    'Argument
    '     asInterval             : 時間間隔を表す文字列
    '                               "yyyy" 年
    '                               "q"    四半期
    '                               "m"    月
    '                               "w"    週
    '                               "d"    日
    '                               "h"    時
    '                               "n"    分
    '                               "s"    秒
    '                               "ms"   ミリ秒
    '     avVal                  : 時間間隔の数値
    '     asSource               : ソース
    'Return Value
    '     指定秒数を加算または減算した新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_addDate( _
        byVal asInterval _
        , byVal avVal _
        , byVal asSource _
        )
        this_addDate = Null
        If this_isInitial() Then Exit Function
        
        '引数の妥当性チェック
        Dim sInterval, boFlg : boFlg = False
        For Each sInterval In Array("yyyy", "q", "m", "w", "d", "h", "n", "s", "ms")
            If cf_isSame(asInterval, sInterval) Then
                boFlg = True : Exit For
            End If
        Next
        ast_argTrue boFlg, asSource, "The value of Interval is invalid."
        ast_argTrue cf_isInteger(avVal), asSource, "The value must be an integer."

        '日時の算出
        Dim dtDateTime
        If cf_isSame(asInterval, "ms") Then
            dtDateTime = DateAdd("s", avVal \ 1000, PdtDateTime)
        Else
            dtDateTime = DateAdd(asInterval, avVal, PdtDateTime)
        End If

        '経過秒の算出
        dbElapsedSeconds = Hour(dtDateTime) * 60 * 60 + Minute(dtDateTime) * 60 + Second(dtDateTime)

        '経過秒の小数部の算出
        Dim dbFractionalPartOfElapsedSeconds : dbFractionalPartOfElapsedSeconds = Null
        If cf_isSame(asInterval, "ms") Then
        '時間間隔の指定がミリ秒の場合
            
            '経過秒の小数部に補正を加える
            dbFractionalPartOfElapsedSeconds = this_getfractionalPartOfElapsedSeconds() + (avVal Mod 1000) / 1000
            
            If dbFractionalPartOfElapsedSeconds >= 1 Then
            '経過秒の小数部が1以上の場合
                '日時と過秒数を+1秒補正
                dtDateTime = DateAdd("s", 1, dtDateTime)
                dbElapsedSeconds = dbElapsedSeconds + 1
                '経過秒の小数部を-1秒補正
                dbFractionalPartOfElapsedSeconds = dbFractionalPartOfElapsedSeconds - 1
            End If
            
            '経過秒の小数部が負の場合は+1秒補正
            If dbFractionalPartOfElapsedSeconds < 0 Then dbFractionalPartOfElapsedSeconds = dbFractionalPartOfElapsedSeconds + 1

        ElseIf Not IsNull(PdbElapsedSeconds) Then
        '時間間隔の指定がミリ秒以外で自身の経過秒がNullでない場合は経過秒の小数部をそのまま使用する
            dbFractionalPartOfElapsedSeconds = this_getfractionalPartOfElapsedSeconds()
        End if
        
        '経過秒の小数部がNullでない（時間間隔の指定がミリ秒か自身の経過秒がNullでない）場合は経過秒を再計算する
        If Not IsNull(dbFractionalPartOfElapsedSeconds) Then dbElapsedSeconds = dbElapsedSeconds + dbFractionalPartOfElapsedSeconds

        '新しいインスタンスを作成して返す
        Set this_addDate = (new Calendar).of(Array(dtDateTime, dbElapsedSeconds))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_correctDatetimeAndElapsedSeconds()
    'Overview                    : 日時と経過秒を補正する
    'Detailed Description        : 経過秒がNullでない場合に、日時の小数部と経過秒の値が整合するよう補正する
    '                              浮動小数点の丸め誤差がある場合は大きい方を採用する
    'Argument
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/11/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_correctDatetimeAndElapsedSeconds( _
        byVal asSource _
        )
        If IsNull(PdbElapsedSeconds) Then Exit Sub

        Dim lFromDateTime, lFromElapsedSeconds
        lFromDateTime = Hour(PdtDateTime) * 60 * 60 + Minute(PdtDateTime) * 60 + Second(PdtDateTime)
        lFromElapsedSeconds = math_tranc(Cdbl(PdbElapsedSeconds))

        '24時間分の秒数の差か1秒以内の差でなければ不整合とみなす
        '浮動小数点の丸め誤差がある場合は大きい方を採用する
        Select Case (lFromDateTime - lFromElapsedSeconds)
        Case 0
            '整合している場合は何もしない
        case (Cl_NUMBER_OF_SECONDS_IN_A_DAY - 1), -1
            PdtDateTime = DateAdd("s", 1, PdtDateTime)
        case (1 - Cl_NUMBER_OF_SECONDS_IN_A_DAY), 1
            PdbElapsedSeconds = Cdbl(Hour(PdtDateTime) * 60 * 60 + Minute(PdtDateTime) * 60 + Second(PdtDateTime))
        Case Else
            ast_failure asSource, "The date/time and elapsed seconds are inconsistent."
        End Select
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_clone()
    'Overview                    : 自身と同じ内容の新しいインスタンスを作る
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     新しいインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_clone( _
        )
        Dim oClone : Set oClone = new Calendar
        If Not this_isInitial() Then oClone.of(Array(PdtDateTime, PdbElapsedSeconds))

        Set this_clone = oClone
        Set oClone = Nothing
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/01         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_compareTo( _
        byRef aoTarget _
        )
        this_compareTo = 0
        If this_isInitial() And IsNull(aoTarget.dateTime) Then Exit Function
        
        Dim lResult : lResult = 0
        If this_isInitial() Or (PdtDateTime < aoTarget.dateTime) Then lResult = -1
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
    'History
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
        If this_isInitial() Then dbResult = -1 * ((aoTarget.dateTime)*Cl_NUMBER_OF_SECONDS_IN_A_DAY + aoTarget.fractionalPartOfElapsedSeconds)
        If IsNull(aoTarget.dateTime) Then dbResult = PdtDateTime*Cl_NUMBER_OF_SECONDS_IN_A_DAY + this_getfractionalPartOfElapsedSeconds
        If dbResult <> 0 Then
            this_differenceFrom = dbResult
            Exit Function
        End If

        Dim dbDiffElapsedSeconds
        dbDiffElapsedSeconds = this_getfractionalPartOfElapsedSeconds-aoTarget.fractionalPartOfElapsedSeconds

        If (PdtDateTime <> aoTarget.dateTime) Then dbDiffElapsedSeconds = dbDiffElapsedSeconds+(PdtDateTime-aoTarget.dateTime)*Cl_NUMBER_OF_SECONDS_IN_A_DAY
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_formatAs( _
        byVal asFormat _
        )
        this_formatAs = "<"&TypeName(Me)&">"&cf_toString(Null)
        If this_isInitial() Then Exit Function

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
    'History
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
    'Function/Sub Name           : this_isInitial()
    'Overview                    : 初期状態かどうかを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     初期状態の場合True、そうでない場合False
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/11/29         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isInitial( _
        )
        this_isInitial = IsNull(PdtDateTime)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_of()
    'Overview                    : 引数に応じたインスタンスを作成する
    'Detailed Description        : this_ofFor1Arg(),this_ofFor2Args()に委譲する
    '                              以下の入力検査を行う
    '                              1.配列でない場合
    '                                this_ofFor1Arg()に委譲する
    '                              2.配列の場合は要素数に応じたチェックを行う
    '                                1-1.要素数が1つ
    '                                 this_ofFor1Arg()に委譲する
    '                                1-2.要素数が2つ
    '                                 this_ofFor2Args()に委譲する
    '                                  e(0) -> 第1引数(Date型)
    '                                  e(1) -> 第2引数(Double型)
    '                                1-3.要素数が6つ
    '                                 this_ofFor2Args()に委譲する
    '                                  e(0-5) -> 第1引数(Date型)
    '                                1-4.要素数が7つ
    '                                 this_ofFor2Args()に委譲する
    '                                  e(0-5) -> 第1引数(Date型)
    '                                  e(6) -> 第2引数(Double型)
    '                                1-5.上記以外の要素数はエラーとする
    'Argument
    '     avArg                  : 引数
    '     asSource               : ソース
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/02         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_of( _
        byRef avArg _
        , byVal asSource _
        )
        ast_argTrue this_isInitial(), asSource, "Because it is an immutable variable, its value cannot be changed."

        If Not(IsArray(avArg)) Then
        '配列でない場合
            this_ofFor1Arg avArg, asSource
        ElseIf new_Arr().hasElement(avArg) Then
        '配列の要素がある場合
            Dim e : e = avArg
            Select Case Ubound(e)
                Case 0:
                    this_ofFor1Arg e(0), asSource
                Case 1:
                    this_ofFor2Args e(0), e(1), asSource
                Case 5:
                    this_ofFor2Args e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5), Null, asSource
                Case 6:
                    this_ofFor2Args e(0)&"/"&e(1)&"/"&e(2)&" "&e(3)&":"&e(4)&":"&e(5), e(6), asSource
                Case Else:
                    ast_failure asSource, "invalid argument. " & cf_toString(avArg)
            End Select
        Else
            ast_failure asSource, "invalid argument. " & cf_toString(avArg)
        End If

        Set this_of = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_ofFor1Arg()
    'Overview                    : 引数が1つの場合のインスタンス作成処理
    'Detailed Description        : 工事中
    'Argument
    '     avDateTime             : 引数の日付時刻
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/02/11         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_ofFor1Arg( _
        byVal avDateTime _
        , byVal asSource _
        )
        Dim oRe : Set oRe = new_Re("^(\s?(?=.*\d)(?:\d{1,4}([-/])\d{1,4}\2\d{1,4})?(?:\s+)?(?:\d{1,2}([-:.])\d{1,2}\3\d{1,2})?)\.([^.]+)$", "")
        If oRe.Test(avDateTime) Then
            Dim dtDateTime : dtDateTime = oRe.Replace(avDateTime, "$1")
            this_setDateTime dtDateTime, asSource
            
            Dim lElapsedSecondsByDt : lElapsedSecondsByDt = Hour(PdtDateTime) * 60 * 60 + Minute(PdtDateTime) * 60 + Second(PdtDateTime)
            this_setElapsedSeconds Cstr(lElapsedSecondsByDt) & "." & oRe.Replace(avDateTime, "$4"), asSource
        Else
            this_setDateTime avDateTime, asSource
        End If
        Set oRe = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_ofFor2Args()
    'Overview                    : 引数が2つの場合のインスタンス作成処理
    'Detailed Description        : 工事中
    'Argument
    '     adtDateTime            : 日時
    '     adbElapsedSeconds      : 経過秒
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_ofFor2Args( _
        byVal adtDateTime _
        , byVal adbElapsedSeconds _
        , byVal asSource _
        )
        this_setDateTime adtDateTime, asSource
        this_setElapsedSeconds adbElapsedSeconds, asSource

        '日時（adtDateTime）と経過秒（adbElapsedSeconds）の整合をチェックする
        this_correctDatetimeAndElapsedSeconds asSource
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setDateTime()
    'Overview                    : 日時（PdtDateTime）のセッター
    'Detailed Description        : 工事中
    'Argument
    '     adtDateTime            : 日時
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setDateTime( _
        byVal adtDateTime _
        , byVal asSource _
        )
        ast_argTrue this_isInitial(), asSource, "Because it is an immutable variable, its value cannot be changed."
        ast_argTrue IsDate(adtDateTime), asSource, "DateTime is not a date/time."
        PdtDateTime = Cdate(adtDateTime)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_setElapsedSeconds()
    'Overview                    : 経過秒（PadbElapsedSeconds）のセッター
    'Detailed Description        : 工事中
    'Argument
    '     adbElapsedSeconds      : 経過秒
    '     asSource               : ソース
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setElapsedSeconds( _
        byVal adbElapsedSeconds _
        , byVal asSource _
        )
        If IsNull(adbElapsedSeconds) Then Exit Sub

        ast_argTrue cf_isNonNegativeNumber(adbElapsedSeconds), asSource, "ElapsedSeconds must be a non-negative number."
        ast_argTrue Cdbl(adbElapsedSeconds) < Cl_NUMBER_OF_SECONDS_IN_A_DAY, asSource, "ElapsedSeconds must be within the number of seconds in a day."

        PdbElapsedSeconds = Cdbl(adbElapsedSeconds)
    End Sub
    
End Class
