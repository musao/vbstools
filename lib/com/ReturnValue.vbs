'***************************************************************************************************
'FILENAME                    : ReturnValue.vbs
'Overview                    : 戻り値クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Class ReturnValue
    'クラス内変数、定数
    Private PvValue, PoErr, PboIsErr
    
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
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PvValue = Empty
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
    'History
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
    'History
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
    'History
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
    'History
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function getErr( _
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function isErr( _
        )
        isErr = PboIsErr
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValue()
    'Overview                    : 戻り値の設定とエラーがあればErrオブジェクトの情報を格納する
    'Detailed Description        : this_setValue()に委譲する
    'Argument
    '     avRet                  : 戻り値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setValue( _
        byRef avRet _
        )
        Set setValue = this_setValue(avRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : setValueByState()
    'Overview                    : 状態による戻り値の設定とエラーがあればErrオブジェクトの情報を格納する
    'Detailed Description        : this_setValueByState()に委譲する
    'Argument
    '     avNormal               : 正常の場合の戻り値
    '     avAbnormal             : 異常の場合の戻り値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function setValueByState( _
        byRef avNormal _
        , byRef avAbnormal _
        )
        Set setValueByState = this_setValueByState(avNormal,avAbnormal)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : toString()
    'Overview                    : オブジェクトの内容を文字列で表示する
    'Detailed Description        : func_CmReturnValueToString()に委譲する
    'Argument
    '     なし
    'Return Value
    '     文字列に変換したオブジェクトの内容
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = func_CmReturnValueToString()
    End Function


    '***************************************************************************************************
    'Function/Sub Name           : this_setValue()
    'Overview                    : 戻り値の設定とエラーがあればErrオブジェクトの情報を格納する
    'Detailed Description        : 工事中
    'Argument
    '     avRet                  : 戻り値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setValue( _
        byRef avRet _
        )
        cf_bind PvValue, avRet
        this_getErrorStatus()
        Set this_setValue = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_setValueByState()
    'Overview                    : 状態による戻り値の設定とエラーがあればErrオブジェクトの情報を格納する
    'Detailed Description        : 工事中
    'Argument
    '     avNormal               : 正常の場合の戻り値
    '     avAbnormal             : 異常の場合の戻り値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_setValueByState( _
        byRef avNormal _
        , byRef avAbnormal _
        )
        If Err.Number=0 Then
            cf_bind PvValue, avNormal
        Else
            cf_bind PvValue, avAbnormal
        End If
        this_getErrorStatus()
        Set this_setValueByState = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_getErrorStatus()
    'Overview                    : エラー状態を取得しエラーがある場合は情報を取得後にクリアする
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/04/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_getErrorStatus( _
        )
        If Err.Number=0 Then
            PboIsErr = False
            Set PoErr = Nothing
        Else
            PboIsErr = True
            Set PoErr = fw_storeErr()
            Err.Clear
        End If
    End Sub
    '***************************************************************************************************
    'Function/Sub Name           : func_CmReturnValueToString()
    'Overview                    : オブジェクトの内容を文字列で表示する
    'Detailed Description        : cf_toString()準拠
    'Argument
    '     なし
    'Return Value
    '     文字列に変換したオブジェクトの内容
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmReturnValueToString( _
        )
        func_CmReturnValueToString = _
            "<" & TypeName(Me) & ">[" _
            & "returnValue:" & cf_toString(PvValue) _
            & ",isErr:" & cf_toString(PboIsErr) _
            & ",getErr:" & cf_toString(PoErr) _
            & "]"
    End Function

End Class
