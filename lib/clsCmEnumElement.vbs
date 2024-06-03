'***************************************************************************************************
'FILENAME                    : clsCmEnumElement.vbs
'Overview                    : Enumの要素クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/26         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmEnumElement
    'クラス内変数、定数
    Private PboAlreadySet, PsKind, PvCode, PsName
    
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
        PsKind = Empty
        PvCode = Empty
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
    'Function/Sub Name           : Property Get code()
    'Overview                    : コード
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     コード
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get code()
        code = func_CmEnumEleGetCode()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get kind()
    'Overview                    : 種類
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     種類
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get kind()
        kind = func_CmEnumEleGetKind()
    End Property

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
        name = func_CmEnumEleGetName()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get toString()
    'Overview                    : インスタンスの内容を文字出力する
    'Detailed Description        : func_CmEnumEleToString()に委譲する
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
        toString = func_CmEnumEleToString()
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : compareTo()
    'Overview                    : 当クラスのインスタンスのcodeを比較する
    'Detailed Description        : func_CmEnumEleCompareTo()に委譲する
    'Argument
    '     aoEnumEle              : 比較対象
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function compareTo( _
        ByRef aoEnumEle _
        )
        Dim vRet : vRet = func_CmEnumEleCompareTo(aoEnumEle)
        ast_argNotNull vRet, TypeName(Me)&"+compareTo()", "The type of the argument is different"
        compareTo = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : equals()
    'Overview                    : 指定されたオブジェクトがこのenum定数と同じ場合にtrueを返す。
    'Detailed Description        : 工事中
    'Argument
    '     aoEnumEle              : 当クラスのインスタンス
    'Return Value
    '     結果 True:一致 / False:不一致
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function equals( _
        ByRef aoEnumEle _
        )
        equals = (func_CmEnumEleCompareTo(aoEnumEle)=0)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : thisIs()
    'Overview                    : 要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     asKind                 : 種類
    '     asName                 : 名前
    '     avCode                 : コード
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function thisIs( _
        ByVal asKind _
        , ByVal asName _
        , ByRef avCode _
        )
        thisIs = Empty
        ast_argFalse PboAlreadySet, TypeName(Me)&"+thisIs()", "Value already set"

        sub_CmEnumEleSetKind asKind
        sub_CmEnumEleSetCode avCode
        sub_CmEnumEleSetName asName
        PboAlreadySet = True
        Set thisIs = Me
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleCompareTo()
    'Overview                    : 当クラスのインスタンスのcodeを比較する
    'Detailed Description        : 下記比較結果を返す
    '                               0  引数と同値
    '                               -1 引数より小さい
    '                               1  引数より大きい
    'Argument
    '     aoEnumEle              : 比較対象
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleCompareTo( _
        ByRef aoEnumEle _
        )
        func_CmEnumEleCompareTo = Null
        If Not cf_isSame(TypeName(Me), TypeName(aoEnumEle)) Then Exit Function
        If Not cf_isSame(PsKind, aoEnumEle.kind) Then Exit Function

        Dim lResult : lResult = 0
        If (PvCode < aoEnumEle.code) Then lResult = -1
        If (PvCode > aoEnumEle.code) Then lResult = 1
        func_CmEnumEleCompareTo = lResult
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetCode()
    'Overview                    : PvCodeのゲッター
    'Detailed Description        : 工事中
    'Argument
    'Return Value
    '     コード
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmEnumEleGetCode( _
        )
        cf_bind func_CmEnumEleGetCode, PvCode
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetKind()
    'Overview                    : PvKindのゲッター
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
    Private Function func_CmEnumEleGetKind( _
        )
        func_CmEnumEleGetKind = PsKind
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleGetName()
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
    Private Function func_CmEnumEleGetName( _
        )
        func_CmEnumEleGetName = PsName
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetCode()
    'Overview                    : PvCodeのセッター
    'Detailed Description        : 工事中
    'Argument
    '     avCode                 : コード
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmEnumEleSetCode( _
        ByVal avCode _
        )
        cf_bind PvCode, avCode
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetKind()
    'Overview                    : PvKindのセッター
    'Detailed Description        : 工事中
    'Argument
    '     asKind                 : 種類
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/05/26         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmEnumEleSetKind( _
        ByVal asKind _
        )
        PsKind = asKind
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmEnumEleSetName()
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
    Private Sub sub_CmEnumEleSetName( _
        ByVal asName _
        )
        PsName = asName
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmEnumEleToString()
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
    Private Function func_CmEnumEleToString( _
        )
        func_CmEnumEleToString = "<" & TypeName(Me) & ">(" & cf_toString(PvCode) & ":" & cf_toString(PsName) & " of " & cf_toString(PsKind) & ")"
    End Function
    
End Class
