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
Class clsCmCalendar
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
       Dim dbFractionalSec : dbFractionalSec = 0
       If Not IsNull(PdbElapsedSeconds) Then dbFractionalSec = PdbElapsedSeconds/(60*60*24)
       serial = Cdbl(PdtDateTime) + dbFractionalSec
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let serial() ->★廃止予定
    'Overview                    : 日付のシリアル値を設定
    'Detailed Description        : シリアル値とは1900/1/1を1として、何日経過したかを示す数値
    'Argument
    '     adbSerial              : 日付のシリアル値
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/17         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let serial( _
        byVal adbSerial _
        )
        Dim dbSec : dbSec = (adbSerial - Fix(adbSerial))*60*60*24
        PdbElapsedSeconds = dbSec - Fix(dbSec)
        PdtDateTime = Cdate(adbSerial - PdbElapsedSeconds/60/60/24)
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
        ast_argsIsSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+compareTo()", "That object is not a calendar class."
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
        ast_argsIsSame TypeName(Me), TypeName(aoTarget), TypeName(Me)&"+differenceFrom()", "That object is not a calendar class."
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
        ByVal asFormat _
        )
        formatAs = this_formatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getNow() ->★廃止予定
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
    Public Function getNow( _
        )
        Set getNow = this_getNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setDateTime() ->★廃止予定
    'Overview                    : 指定した日付時刻を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avDateTime             : 設定する日付時刻
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setDateTime( _
        ByVal avDateTime _
        )
        Set setDateTime = this_setDate(avDateTime)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : newInstance()
    'Overview                    : インスタンスを作成する
    'Detailed Description        : this_newInstance()に委譲する
    'Argument
    '     adtDateTime            : 日時
    '     adbElapsedSeconds      : 経過秒
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/09/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function newInstance( _
        ByVal adtDateTime _
        , ByVal adbElapsedSeconds _
        )
        Set newInstance = this_newInstance(adtDateTime, adbElapsedSeconds, TypeName(Me)&"+newInstance()")
    End Function
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getNow() ->★廃止予定
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
    Private Function this_getNow( _
        )
        PdtDateTime = Now()
        
        Dim dbTimer : dbTimer = Timer()
        PdbElapsedSeconds = dbTimer - Fix(dbTimer)

        Set this_getNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setDate() ->★廃止予定
    'Overview                    : 指定した日付時刻を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avDateTime             : 設定する日付時刻
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setDate( _
        ByVal avDateTime _
        )
        Dim sPtn : sPtn = "^([^.]+)\.(\d+)$"
        If new_Re(sPtn, "").Test(avDateTime) Then
            PdtDateTime = Cdate(new_Re(sPtn, "").Replace(avDateTime, "$1"))
            PdbElapsedSeconds = Cdbl("0." & new_Re(sPtn, "").Replace(avDateTime, "$2"))
        Else
            PdtDateTime = Cdate(avDateTime)
            PdbElapsedSeconds = Null
        End If
        Set this_setDate = Me
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
                    
                    If StrComp(sKey, Mid(asFormat, lPos, lKeyLen))=0 Then
                    '変換テーブルにある文字と一致した場合
                        vItem = .Item(sKey)
                        If cf_isSame(Cl_USE_DATAPART, vItem(0)) Then
                        'PdtDateからDatePart()で値を取り出す場合
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDateTime), lKeyLen, "0", vItem(2), True)
                        Else
                        '秒数の小数部を取り出す場合
                            Dim dbFractionalSec : dbFractionalSec =0
                            If Not IsNull(PdbElapsedSeconds) Then dbFractionalSec = PdbElapsedSeconds
                            sItemValue = func_CM_FillInTheCharacters(Fix(dbFractionalSec*10^lKeyLen), lKeyLen, "0", False, True)
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
    'Function/Sub Name           : this_newInstance()
    'Overview                    : インスタンスを作成する
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
    Private Function this_newInstance( _
        byVal adtDateTime _
        , byVal adbElapsedSeconds _
        , byVal asSource _
        )
        ast_argFalse IsNull(PdtDateTime), asSource, "Because it is an immutable variable, its value cannot be changed."
        this_setDateTime adtDateTime, asSource
        this_setElapsedSeconds adbElapsedSeconds, asSource
        Set this_newInstance = Me
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
        ast_argTrue IsDate(PdtDateTime), asSource, "DateTime is not a date/time."
        PadtDateTime = adtDateTime
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
        PadbElapsedSeconds = adbElapsedSeconds
    End Sub
    
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
        Dim lResult : lResult = 0

        If (Me.dateTime < aoTarget.dateTime) Then lResult = -1
        If (Me.dateTime > aoTarget.dateTime) Then lResult = 1
        If lResult <> 0 Then
            this_compareTo = lResult
            Exit Function
        End If
        
        If (Me.elapsedSeconds < aoTarget.elapsedSeconds) Then lResult = -1
        If (Me.elapsedSeconds > aoTarget.elapsedSeconds) Then lResult = 1

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

        Dim dbDiffElapsedSeconds : dbDiffElapsedSeconds = Me.elapsedSeconds - aoTarget.elapsedSeconds
        If (Me.dateTime <> aoTarget.dateTime) Then dbDiffElapsedSeconds = dbDiffElapsedSeconds+(Me.dateTime-aoTarget.dateTime)*60*60*24
        this_differenceFrom = math_roundDown(dbDiffElapsedSeconds, 5)

'        this_differenceFrom = math_roundDown(Me.serial()*60*60*24-aoTarget.serial()*60*60*24, 5)
    End Function
    
End Class
