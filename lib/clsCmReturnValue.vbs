'***************************************************************************************************
'FILENAME                    : clsCmReturnValue.vbs
'Overview                    : 戻り値クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmReturnValue
    'クラス内変数、定数
    Private PvValue
    Private PoErr
    Private PboIsErr
    
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
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PvValue = Nothing
        Set PoErr = Nothing
        PboIsErr = Empty
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
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PvValue = Nothing
        Set PoErr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get returnValue()
    'Overview                    : 戻り値を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     戻り値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get returnValue()
        cf_bind returnValue, PvValue
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let returnValue()
    'Overview                    : 戻り値を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avRet                  : 戻り値
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let returnValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set returnValue()
    'Overview                    : 戻り値を設定する
    'Detailed Description        : 工事中
    'Argument
    '     avRet                  : 戻り値
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set returnValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : getErr()
    'Overview                    : エラー情報を返却する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     エラー情報
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function getErr( _
        )
        Set getErr = PoErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : isErr()
    'Overview                    : エラー情報の有無を返却する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     結果 True:エラーあり / False:エラーなし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function isErr( _
        )
        isErr = PboIsErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValue()
    'Overview                    : 戻り値の取得とエラーがあればErrオブジェクトの情報を格納する
    'Detailed Description        : 工事中
    'Argument
    '     avRet                  : 戻り値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Function setValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
        If Err.Number=0 Then
            PboIsErr = False
            Set PoErr = Nothing
        Else
            PboIsErr = True
            Set PoErr = fw_storeErr()
        End If
        Set setValue = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : toString()
    'Overview                    : オブジェクトの内容を文字列で表示する
    'Detailed Description        : cf_toString()準拠
    'Argument
    '     なし
    'Return Value
    '     文字列に変換したオブジェクトの内容
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = _
            "<" & TypeName(Me) & ">[" _
            & "returnValue:" & cf_toString(PvValue) _
            & ",isErr:" & cf_toString(PboIsErr) _
            & ",getErr:" & cf_toString(PoErr) _
            & "]"
    End Function

End Class
