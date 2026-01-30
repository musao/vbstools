'***************************************************************************************************
'FILENAME                    : HtmlGenerator.vbs
'Overview                    : HTML生成クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/22         Y.Fujii                  First edition
'***************************************************************************************************
Class HtmlGenerator
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
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTagInfo = new_DicOf(Array("element", Empty, "attribute", Empty, "content", Empty))
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
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoTagInfo = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get attribute()
    'Overview                    : 属性（オブジェクトの配列）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     属性（オブジェクトの配列）を返す
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get attribute()
        attribute = PoTagInfo.Item("attribute")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get content()
    'Overview                    : 内容（オブジェクトの配列）を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     内容（オブジェクトの配列）を返す
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get content()
        content = PoTagInfo.Item("content")
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let element()
    'Overview                    : 要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     asElement              : 要素
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let element( _
        byVal asElement _
        )
        PoTagInfo.Item("element") = asElement
'        If new_Re("^[!-~][ -~]*$", "i").Test(asElement) Then
'            PoTagInfo.Item("element") = asElement
'        Else
'            Err.Raise 1032, "HtmlGenerator.vbs:HtmlGenerator+element()", "要素（element）には半角以外の文字を指定できません。"
'        End If
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get element()
    'Overview                    : 要素を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     要素
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get element()
        element = PoTagInfo.Item("element")
    End Property
        
    '***************************************************************************************************
    'Function/Sub Name           : addContent()
    'Overview                    : 内容を追加する
    'Detailed Description        : 工事中
    'Argument
    '     avCont                 : 追加する内容
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addContent( _
        byRef avCont _
        )
        Dim vArr : cf_bind vArr, PoTagInfo.Item("content")
        cf_push vArr, avCont
        cf_bindAt PoTagInfo, "content", vArr

        Set addContent = Me
    End Function
        
    '***************************************************************************************************
    'Function/Sub Name           : addAttribute()
    'Overview                    : 属性を追加する
    'Detailed Description        : 工事中
    'Argument
    '     asKey                  : 追加する属性のキー
    '     asValue                : 追加する属性の値
    'Return Value
    '     自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function addAttribute( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicOf(Array("key", asKey, "value", asValue))
        Dim vArr : cf_bind vArr, PoTagInfo.Item("attribute")
        cf_push vArr, oNewAttr
        cf_bindAt PoTagInfo, "attribute", vArr

        Set addAttribute = Me
        Set oNewAttr = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : HTMLを生成する
    'Detailed Description        : this_generate()に委譲する
    'Argument
    '     なし
    'Return Value
    '     生成したHTML
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
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
    'Overview                    : HTMLを生成する
    'Detailed Description        : 工事中
    'Argument
    '     asSource               : ソース
    'Return Value
    '     生成したHTML
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_generate( _
        byVal asSource _
        )
        ast_argNotEmpty PoTagInfo.Item("element"), asSource, "HTML tags without elements cannot be generated."
'        If IsEmpty(PoTagInfo.Item("element")) Then
'            Err.Raise 17, "HtmlGenerator.vbs:HtmlGenerator-this_generate()", "要素がないHTMLタグは生成できません。"
'            Exit Function
'        End If

        '開始タグの編集
        Dim sStt : sStt =  "<" & PoTagInfo.Item("element")
        Dim vArr, vEle
        '属性（attribute）の編集
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attributeが空でない場合
            For Each vEle In PoTagInfo.Item("attribute")
                cf_push vArr, this_editAttribute(vEle)
            Next
            sStt = sStt & " " & Join(vArr, " ")
        End If
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            sStt = sStt & ">"
        Else
        'contentが空の場合
            sStt = sStt & " />"
        End If
        
        '内容（content）の編集
        Dim sCont : sCont = ""
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            vArr = Array()
            For Each vEle In PoTagInfo.Item("content")
                cf_push vArr, this_editContent(vEle)
            Next
            sCont = new_Re("^([^\n])", "igm").Replace(Join(vArr, vbNewLine),"  $1")
        End If

        '終了タグの編集
        Dim sEnd : sEnd =  ""
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            sEnd =  "</" & PoTagInfo.Item("element") & ">"
        End If

        '生成したHTMLを返却
        sRet = sStt
        If Not IsEmpty(PoTagInfo.Item("content")) Then sRet = sRet & vbNewLine & sCont & vbNewLine & sEnd
        this_generate = sRet

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editAttribute()
    'Overview                    : 属性（attribute）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoAttr                 : 編集する属性（attribute）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_editAttribute( _
        byRef aoAttr _
        )
        Dim sRet
        If IsEmpty(aoAttr.Item("value")) Then
            sRet = aoAttr.Item("key")
        Else
            sRet = aoAttr.Item("key") & "=" & Chr(34) & aoAttr.Item("value") & Chr(34)
        End If
        this_editAttribute = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_editContent()
    'Overview                    : 内容（content）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoCont                 : 編集する内容（content）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_editContent( _
        byRef aoCont _
        )
        Dim sRet
        On Error Resume Next
        sRet = aoCont.generate()
        If Err.Number<>0 Then
            sRet = this_htmlEntityReference(aoCont)
        End If
        On Error GoTo 0
        this_editContent = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : this_htmlEntityReference()
    'Overview                    : HTMLの特殊文字を実体参照（entity reference）処理する
    'Detailed Description        : HTMLとして特殊な意味を持つ文字（特殊文字またはメタ文字）を意味を持たない
    '                              別の文字列に置換する
    'Argument
    '     asTarget               : 実体参照処理する文字列
    'Return Value
    '     実体参照処理した文字列
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function this_htmlEntityReference( _
        byRef asTarget _
        )
        Dim vSettings
        Select Case LCase(PoTagInfo("element"))
        Case "script"
            vSettings = Array( _
                Array("</script>", "<\/script>") _
                )
        Case "style"
            vSettings = Array( _
                Array("</style>", "<\/style>") _
                )
        Case "textarea"
            vSettings = Array( _
                Array("</textarea>", "<\/textarea>") _
                )
        Case Else
            vSettings = Array( _
                Array("&", "&amp;") _
                , Array("'", "&#39;") _
                , Array("""", "&quot;") _
                , Array("<", "&lt;") _
                , Array(">", "&gt;") _
                )
        End Select
        Dim sTarget : sTarget = asTarget
        Dim i
        For Each i In vSettings
            sTarget = Replace(sTarget, i(0), i(1))
        Next
        this_htmlEntityReference = sTarget
    End Function

End Class
