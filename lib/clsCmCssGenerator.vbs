'***************************************************************************************************
'FILENAME                    : clsCmCssGenerator.vbs
'Overview                    : CSS生成クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/25         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmCssGenerator
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
    'Histroy
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
    'Histroy
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
    'Histroy
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
    'Histroy
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
'            Err.Raise 1032, "clsCmCssGenerator.vbs:clsCmCssGenerator+selector()", "セレクタには半角以外の文字を指定できません。"
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
    'Histroy
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
    'Histroy
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
    'Detailed Description        : func_CmCssGenGenerate()に委譲する
    'Argument
    '     なし
    'Return Value
    '     生成したCSS
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function generate( _
        )
        generate = func_CmCssGenGenerate()
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
    '2023/12/27         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        toString = func_CmCssGenGenerate()
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : func_CmCssGenGenerate()
    'Overview                    : CSSを生成する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     生成したCSS
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCssGenGenerate( _
        )
        If IsEmpty(PoTagInfo.Item("selector")) Then
            Err.Raise 17, "clsCmCssGenerator.vbs:clsCmCssGenerator-func_CmCssGenGenerate()", "セレクタがないCSSは生成できません。"
            Exit Function
        End If

        Dim sRet : sRet = PoTagInfo.Item("selector") & " {" & vbNewLine

        'プロパティ（property）の編集
        Dim vArr, vEle
        If Not IsEmpty(PoTagInfo.Item("property")) Then
        'propertyが空でない場合
            For Each vEle In PoTagInfo.Item("property")
                cf_push vArr, "  " & func_CmCssGenEditProperty(vEle)
            Next
            sRet = sRet & Join(vArr, vbNewLine) & vbNewLine
        End If
        
        sRet = sRet & "}"

        '生成したCSSを返却
        func_CmCssGenGenerate = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmCssGenEditProperty()
    'Overview                    : プロパティ（property）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoAttr                 : 編集するプロパティ（property）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/25         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmCssGenEditProperty( _
        byRef aoAttr _
        )
        func_CmCssGenEditProperty = aoAttr.Item("key") & " : " & aoAttr.Item("value") & " ;"
    End Function

End Class
