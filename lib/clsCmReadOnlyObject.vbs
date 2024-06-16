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
        name = func_CmReadOnlyObjectGetName()
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
        cf_bind parent, func_CmReadOnlyObjectGetParent()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : インスタンスの内容を文字出力する
    'Detailed Description        : func_CmReadOnlyObjectToString()に委譲する
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
        toString = func_CmReadOnlyObjectToString()
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
        cf_bind value, func_CmReadOnlyObjectGetValue()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : 当クラスのインスタンスのvalueを比較する
    'Detailed Description        : func_CmReadOnlyObjectCompareTo()に委譲する
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
        Dim vRet : vRet = func_CmReadOnlyObjectCompareTo(aoTarget)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different"
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
        equals = (func_CmReadOnlyObjectCompareTo(aoTarget)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : thisIs()
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
    Public Function thisIs( _
        ByRef aoParent _
        , ByVal asName _
        , ByRef avValue _
        )
        thisIs = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+thisIs()", "Value already set"

        sub_CmReadOnlyObjectSetParent aoParent
        sub_CmReadOnlyObjectSetName asName
        sub_CmReadOnlyObjectSetValue avValue
        PboAlreadySet = True
        Set thisIs = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : func_CmReadOnlyObjectCompareTo()
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
    Private Function func_CmReadOnlyObjectCompareTo( _
        ByRef aoTarget _
        )
        func_CmReadOnlyObjectCompareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoTarget)) Then Exit Function
        If Not cf_isSame(PoParent, aoTarget.parent) Then Exit Function

        Dim lResult : lResult = 0
        If (PvValue < aoTarget.value) Then lResult = -1
        If (PvValue > aoTarget.value) Then lResult = 1
        func_CmReadOnlyObjectCompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmReadOnlyObjectGetValue()
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
    Private Function func_CmReadOnlyObjectGetValue( _
        )
        cf_bind func_CmReadOnlyObjectGetValue, PvValue
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmReadOnlyObjectGetName()
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
    Private Function func_CmReadOnlyObjectGetName( _
        )
        func_CmReadOnlyObjectGetName = PsName
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmReadOnlyObjectGetParent()
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
    Private Function func_CmReadOnlyObjectGetParent( _
        )
        cf_bind func_CmReadOnlyObjectGetParent, PoParent
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmReadOnlyObjectSetName()
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
    Private Sub sub_CmReadOnlyObjectSetName( _
        ByVal asName _
        )
        PsName = asName
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmReadOnlyObjectSetParent()
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
    Private Sub sub_CmReadOnlyObjectSetParent( _
        ByVal aoParent _
        )
        cf_bind PoParent, aoParent
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmReadOnlyObjectSetValue()
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
    Private Sub sub_CmReadOnlyObjectSetValue( _
        ByRef avValue _
        )
        cf_bind PvValue, avValue
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmReadOnlyObjectToString()
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
    Private Function func_CmReadOnlyObjectToString( _
        )
        func_CmReadOnlyObjectToString = "<" & TypeName(Me) & ">(" & cf_toString(PvValue) & ":" & cf_toString(PsName) & " of " & cf_toString(PoParent) & ")"
    End Function
    
End Class
