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
    Private PdtDate, PdbFractionalSec, PsDefaultFormat
    
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
        PdtDate = Empty
        PdbFractionalSec = Empty
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
       If Not IsEmpty(PdbFractionalSec) Then dbFractionalSec = PdbFractionalSec/(60*60*24)
       serial = Cdbl(PdtDate) + dbFractionalSec
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let serial()
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
        PdbFractionalSec = dbSec - Fix(dbSec)
        PdtDate = Cdate(adbSerial - PdbFractionalSec/60/60/24)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : デフォルトの形式で表示する
    'Detailed Description        : func_CmCalendaFormatAs()に委譲する
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
        toString = func_CmCalendaFormatAs(PsDefaultFormat)
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
        If Strcomp(TypeName(Me), TypeName(aoTarget), vbBinaryCompare)<>0 Then
            Err.Raise 438, "clsCmCalendar.vbs:clsCmCalendar+compareTo()", "オブジェクトでサポートされていないプロパティまたはメソッドです。"
            Exit Function
        End If

        Dim SerialMe : SerialMe = Me.serial()
        Dim SerialTg : SerialTg = aoTarget.serial()
        Dim lResult : lResult = 0
        If (SerialMe < SerialTg) Then lResult = -1
        If (SerialMe > SerialTg) Then lResult = 1
        compareTo = lResult
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
        If Strcomp(TypeName(Me), TypeName(aoTarget), vbBinaryCompare)<>0 Then
            Err.Raise 438, "clsCmCalendar.vbs:clsCmCalendar+differenceFrom()", "オブジェクトでサポートされていないプロパティまたはメソッドです。"
            Exit Function
        End If

        differenceFrom = math_roundDown(Me.serial()*60*60*24-aoTarget.serial()*60*60*24, 5)

    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : formatAs()
    'Overview                    : 日付を整形する
    'Detailed Description        : func_CmCalendaFormatAs()に委譲する
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
        formatAs = func_CmCalendaFormatAs(asFormat)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getNow()
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
        Set getNow = func_CmCalendarGetNow()
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : setDateTime()
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
        Set setDateTime = func_CmCalendarSetDate(avDateTime)
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
        PdtDate = Now()
        
        Dim dbTimer : dbTimer = Timer()
        PdbFractionalSec = dbTimer - Fix(dbTimer)

        Set func_CmCalendarGetNow = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendarSetDate()
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
    Private Function func_CmCalendarSetDate( _
        ByVal avDateTime _
        )
        Dim sPtn : sPtn = "^([^.]+)\.(\d+)$"
        If new_Re(sPtn, "").Test(avDateTime) Then
            PdtDate = Cdate(new_Re(sPtn, "").Replace(avDateTime, "$1"))
            PdbFractionalSec = Cdbl("0." & new_Re(sPtn, "").Replace(avDateTime, "$2"))
        Else
            PdtDate = Cdate(avDateTime)
            PdbFractionalSec = Empty
        End If
        Set func_CmCalendarSetDate = Me
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmCalendaFormatAs()
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
    Private Function func_CmCalendaFormatAs( _
        byVal asFormat _
        )
        Dim oConversionSettings : Set oConversionSettings = new_Dic()
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
            .Add "000000", Array("GetFractionalSeconds")
            .Add "000", Array("GetFractionalSeconds")
            
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
                        'PdtDateからDatePart()で値を取り出す場合
                            sItemValue = func_CM_FillInTheCharacters(DatePart(vItem(1), PdtDate), lKeyLen, "0", vItem(2), True)
                        Else
                        '秒数の小数部を取り出す場合
                            Dim dbFractionalSec : dbFractionalSec =0
                            If Not IsEmpty(PdbFractionalSec) Then dbFractionalSec = PdbFractionalSec
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
        func_CmCalendaFormatAs = sResult
        Set oConversionSettings = Nothing
    End Function
    
End Class
