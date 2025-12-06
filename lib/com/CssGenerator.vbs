'***************************************************************************************************
'FILENAME                    : CssGenerator.vbs
'Overview                    : CSS生成クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/25         Y.Fujii                  First edition
'***************************************************************************************************
Class CssGenerator
    'クラス内変数、定数
    Private PoTagInfo
    
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
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTagInfo = new_DicOf(Array("selector", Empty, "property", Empty))
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
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoTagInfo = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get property()
    'Overview                    : プロパティ（オブジェクトの配列）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     プロパティ（オブジェクトの配列）を返す
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get property()
        property = PoTagInfo.Item("property")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let selector()
    'Overview                    : セレクタを設定する
    'Detailed Description        : 工事中
    'Argument
    '     asSelector             : セレクタ
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let selector( _
        byVal asSelector _
        )
        PoTagInfo.Item("selector") = asSelector
'        If new_Re("^[!-~][ -~]*$", "i").Test(asSelector) Then
'            PoTagInfo.Item("selector") = asSelector
'        Else
'            Err.Raise 1032, "CssGenerator.vbs:CssGenerator+selector()", "セレクタには半角以外の文字を指定できません。"
'        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get selector()
    'Overview                    : セレクタを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     セレクタ
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get selector()
        selector = PoTagInfo.Item("selector")
    End Property
        
    '***************************************************************************************************
    'Function/Sub Name           : addProperty()
    'Overview                    : プロパティを追加する
    'Detailed Description        : 工事中
    'Argument
    '     asKey                  : 追加するプロパティのキー
    '     asValue                : 追加するプロパティの値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addProperty( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicOf(Array("key", asKey, "value", asValue))
        Dim vArr : cf_bind vArr, PoTagInfo.Item("property")
        cf_push vArr, oNewAttr
        cf_bindAt PoTagInfo, "property", vArr

        Set addProperty = Me
        Set oNewAttr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : CSSを生成する
    'Detailed Description        : this_generate()に委譲する
    'Argument
    '     なし
    'Return Value
    '     生成したCSS
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function generate( _
        )
        generate = this_generate(TypeName(Me)&"+generate()")
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = this_generate(TypeName(Me)&"+toString()")
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : this_generate()
    'Overview                    : CSSを生成する
    'Detailed Description        : 工事中
    'Argument
    '     asSource               : ソース
    'Return Value
    '     生成したCSS
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_generate( _
        byVal asSource _
        )
        ast_argNotEmpty PoTagInfo.Item("selector"), asSource, "CSS without selectors cannot be generated."

        Dim sRet : sRet = PoTagInfo.Item("selector") & " {" & vbNewLine

        'プロパティ（property）の編集
        Dim vArr, vEle
        If Not IsEmpty(PoTagInfo.Item("property")) Then
        'propertyが空でない場合
            For Each vEle In PoTagInfo.Item("property")
                cf_push vArr, "  " & this_editProperty(vEle)
            Next
            sRet = sRet & Join(vArr, vbNewLine) & vbNewLine
        End If
        
        sRet = sRet & "}"

        '生成したCSSを返却
        this_generate = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editProperty()
    'Overview                    : プロパティ（property）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoAttr                 : 編集するプロパティ（property）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_editProperty( _
        byRef aoAttr _
        )
        this_editProperty = aoAttr.Item("key") & " : " & aoAttr.Item("value") & " ;"
    End Function

End Class
