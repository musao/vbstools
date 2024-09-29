'***************************************************************************************************
'FILENAME                    : clsCmReadOnlyObject.vbs
'Overview                    : 読み取り専用オブジェクトのクラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmReadOnlyObject
    'クラス内変数、定数
    Private PboAlreadySet, PoParent, PvValue, PsName
    
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
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        PboAlreadySet = False
        Set PoParent = Nothing
        PvValue = Empty
        PsName = Empty
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
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoParent = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : Property Get name()
    'Overview                    : 名前
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     名前
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get name()
        name = this_getName()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get parent()
    'Overview                    : 親
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     親
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get parent()
        cf_bind parent, this_getParent()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : インスタンスの内容を文字出力する
    'Detailed Description        : this_toString()に委譲する
    'Argument
    '     なし
    'Return Value
    '     インスタンスの内容
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get toString()
        toString = this_toString()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get value()
    'Overview                    : 値
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get value()
        cf_bind value, this_getValue()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : 当クラスのインスタンスのvalueを比較する
    'Detailed Description        : this_compareTo()に委譲する
    'Argument
    '     aoTarget               : 比較対象
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function compareTo( _
        ByRef aoTarget _
        )
        Dim vRet : vRet = this_compareTo(aoTarget)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different."
        compareTo = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : equals()
    'Overview                    : 指定されたオブジェクトがこのenum定数と同じ場合にtrueを返す。
    'Detailed Description        : 工事中
    'Argument
    '     aoTarget               : 当クラスのインスタンス
    'Return Value
    '     結果 True:一致 / False:不一致
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function equals( _
        ByRef aoTarget _
        )
        equals = (this_compareTo(aoTarget)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : is()
    'Overview                    : 値を設定する
    'Detailed Description        : 既に設定済みの場合は例外
    'Argument
    '     aoParent               : 親のオブジェクト
    '     asName                 : 名前
    '     avValue                : 値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function [is]( _
        ByRef aoParent _
        , ByVal asName _
        , ByRef avValue _
        )
        [is] = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+is()", "Value already set."

        this_setParent aoParent
        this_setName asName
        this_setValue avValue
        PboAlreadySet = True
        Set [is] = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : this_compareTo()
    'Overview                    : 当クラスのインスタンスのvalueを比較する
    'Detailed Description        : 下記比較結果を返す
    '                               0  引数と同値
    '                               -1 引数より小さい
    '                               1  引数より大きい
    'Argument
    '     aoTarget              : 比較対象
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_compareTo( _
        ByRef aoTarget _
        )
        this_compareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoTarget)) Then Exit Function
        If Not cf_isSame(PoParent, aoTarget.parent) Then Exit Function

        Dim lResult : lResult = 0
        If (PvValue < aoTarget.value) Then lResult = -1
        If (PvValue > aoTarget.value) Then lResult = 1
        this_compareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getValue()
    'Overview                    : PvValueのゲッター
    'Detailed Description        : 工事中
    'Argument
    'Return Value
    '     値
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getValue( _
        )
        cf_bind this_getValue, PvValue
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getName()
    'Overview                    : PsNameのゲッター
    'Detailed Description        : 工事中
    'Argument
    'Return Value
    '     名前
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getName( _
        )
        this_getName = PsName
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_getParent()
    'Overview                    : PvParentのゲッター
    'Detailed Description        : 工事中
    'Argument
    'Return Value
    '     種類
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_getParent( _
        )
        cf_bind this_getParent, PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setName()
    'Overview                    : PsNameのセッター
    'Detailed Description        : 工事中
    'Argument
    '     asName                 : 名前
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setName( _
        ByVal asName _
        )
        PsName = asName
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setParent()
    'Overview                    : PvParentのセッター
    'Detailed Description        : 工事中
    'Argument
    '     aoParent               : 親
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setParent( _
        ByVal aoParent _
        )
        cf_bind PoParent, aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : this_setValue()
    'Overview                    : PvValueのセッター
    'Detailed Description        : 工事中
    'Argument
    '     avValue                : 値
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_setValue( _
        ByRef avValue _
        )
        cf_bind PvValue, avValue
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : this_toString()
    'Overview                    : インスタンスの内容を文字出力する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     インスタンスの内容
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_toString( _
        )
        this_toString = "<" & TypeName(Me) & ">{" & cf_toString(PsName) & ":" & cf_toString(PvValue) & "}"
    End Function
    
End Class
