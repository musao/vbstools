'***************************************************************************************************
'FILENAME                    : Cash.vbs
'Overview                    : 汎用キャッシュクラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/12/06         Y.Fujii                  First edition
'***************************************************************************************************
Class Cash
    Private PoCasher

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
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoCasher = new_Dic()
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
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoCasher = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get item()
    'Overview                    : 値を取得する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get item( _
        byVal avKey _
        )
        cf_bind item, this_get(avKey)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : デフォルトの形式で表示する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     インスタンスの内容を文字列で表現したもの
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = "<"&TypeName(Me)&">"&Mid(cf_toString(PoCasher), Len("<Dictionary>")+1)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : clear()
    'Overview                    : キャッシュを全てクリアする
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub clear( _
        )
        PoCasher.RemoveAll
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : delete()
    'Overview                    : 値を削除する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub delete( _
        byVal avKey _
        )
        this_delete avKey
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : get()
    'Overview                    : 値を取得する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function [get]( _
        byVal avKey _
        )
        cf_bind [get], this_get(avKey)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : getOrCompute()
    'Overview                    : キャッシュにあれば値を返し、なければ生成関数を実行して値を取得し
    '                              その値を返しつつキャッシュに保存する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    '     aoLoader               : 値を生成する関数
    '     alTtl                  : 生存期間（ミリ秒）
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function getOrCompute( _
        byVal avKey _
        , byVal aoLoader _
        , byVal alTtl _
        )
        If Not this_has(avKey) Then this_put avKey, aoLoader(avKey), alTtl
        cf_bind getOrCompute, this_get(avKey)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : has()
    'Overview                    : キーが存在するか確認する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     結果 True:キーが存在する / False:キーが存在しない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function has( _
        byVal avKey _
        )
        has = this_has(avKey)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : put()
    'Overview                    : 値を設定する
    'Detailed Description        : 生存期間（ミリ秒）が負値の場合は無期限とする
    '                              this_put()に委譲する
    'Argument
    '     avKey                  : キー
    '     avValue                : 値
    '     alTtl                  : 生存期間（ミリ秒）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub put( _
        byVal avKey _
        , byVal avValue _
        , byVal alTtl _
        )
        this_put avKey, avValue, alTtl
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : size()
    'Overview                    : 有効なキャッシュのサイズを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     キャッシュのサイズ
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function size( _
        )
        size = this_size()
    End Function






    '***************************************************************************************************
    'Function/Sub Name           : this_delete()
    'Overview                    : 値を削除する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_delete( _
        byVal avKey _
        )
        If PoCasher.Exists(avKey) Then PoCasher.Remove(avKey)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_get()
    'Overview                    : 値を取得する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_get( _
        byVal avKey _
        )
        this_get = Null
        If this_has(avKey) Then cf_bind this_get, PoCasher(avKey).Item("value")
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_has()
    'Overview                    : 有効期間内のキーが存在するか確認する
    'Detailed Description        : 工事中
    'Argument
    '     avKey                  : キー
    'Return Value
    '     結果 True:キーが存在する / False:キーが存在しない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_has( _
        byVal avKey _
        )
        this_has = False
        If Not PoCasher.Exists(avKey) Then Exit Function

        Dim boInTime : boInTime = this_isInTime(new_Now(), avKey)
        If Not boInTime Then this_delete(avKey)
        this_has = boInTime
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_isInTime()
    'Overview                    : キーが有効期間内か確認する
    'Detailed Description        : 工事中
    'Argument
    '     aoRefTime              : 判定基準時間
    '     avKey                  : キー
    'Return Value
    '     結果 True:キーが有効期間内 / False:キーが有効期間内でない
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_isInTime( _
        byRef aoRefTime _
        , byVal avKey _
        )
        this_isInTime = True
        If Not IsNull(PoCasher(avKey).Item("time")) Then this_isInTime = Not (aoRefTime.compareTo( PoCasher(avKey).Item("time") ) > 0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_put()
    'Overview                    : 値を設定する
    'Detailed Description        : 生存期間（ミリ秒）が負値の場合は無期限とする
    'Argument
    '     avKey                  : キー
    '     avValue                : 値
    '     alTtl                  : 生存期間（ミリ秒）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_put( _
        byVal avKey _
        , byVal avValue _
        , byVal alTtl _
        )
        Dim oTime : oTime = Null
        If alTtl >= 0 Then Set oTime = new_Now().addMilliseconds(alTtl)

        cf_bindAt PoCasher, avKey, _
                                new_DicOf( _
                                  Array( "value", avValue _
                                        , "time" , oTime _
                                        ) _
                                 )
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_size()
    'Overview                    : 有効なキャッシュのサイズを取得する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     キャッシュのサイズ
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2025/12/07         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_size( _
        )
        Dim vKey, oNow : Set oNow = new_Now()
        
        For Each vKey In PoCasher.Keys
            If Not this_isInTime(oNow, vKey) Then this_delete(vKey)
        Next
        this_size = PoCasher.Count
        
        Set oNow = Nothing
    End Function

End Class
