'***************************************************************************************************
'FILENAME                    : clsCmHtmlGenerator.vbs
'Overview                    : HTML生成クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/22         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmHtmlGenerator
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
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoTagInfo = new_DicWith(Array("element", Empty, "attribute", Empty, "content", Empty))
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
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get attribute()
        attribute = PoTagInfo.Item("attribute").Items()
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
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get content()
        content = PoTagInfo.Item("content").Items()
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
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let element( _
        byVal asElement _
        )
        If new_Re("^[!-~][ -~]*$", "i").Test(asElement) Then
            PoTagInfo.Item("element") = asElement
        Else
            Err.Raise 1032, "clsCmHtmlGenerator.vbs:clsCmHtmlGenerator+element()", "要素（element）には半角以外の文字を指定できません。"
        End If
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
    'Histroy
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
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub addContent( _
        byRef avCont _
        )
        If IsEmpty(PoTagInfo.Item("content")) Then
            Set PoTagInfo.Item("content") = new_ArrWith(avCont)
        Else
            PoTagInfo.Item("content").push avCont
        End If
        Set oNewAttr = Nothing
    End Sub
        
    '***************************************************************************************************
    'Function/Sub Name           : addAttribute()
    'Overview                    : 属性を追加する
    'Detailed Description        : 工事中
    'Argument
    '     asKey                  : 追加する属性のキー
    '     asValue                : 追加する属性の値
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub addAttribute( _
        byVal asKey _
        , byVal asValue _
        )
        Dim oNewAttr : Set oNewAttr = new_DicWith(Array("key", asKey, "value", asValue))
        If IsEmpty(PoTagInfo.Item("attribute")) Then
            Set PoTagInfo.Item("attribute") = new_ArrWith(oNewAttr)
        Else
            PoTagInfo.Item("attribute").push oNewAttr
        End If
        Set oNewAttr = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : generate()
    'Overview                    : HTMLを生成する
    'Detailed Description        : func_CmHtmlGenGenerate()に委譲する
    'Argument
    '     なし
    'Return Value
    '     生成したHTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function generate( _
        )
        generate = func_CmHtmlGenGenerate()
    End Function




    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenGenerate()
    'Overview                    : HTMLを生成する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     生成したHTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenGenerate( _
        )
        If IsEmpty(PoTagInfo.Item("element")) Then
            Err.Raise 17, "clsCmHtmlGenerator.vbs:clsCmHtmlGenerator-func_CmHtmlGenGenerate()", "要素がないHTMLタグは生成できません。"
            Exit Function
        End If

        Dim sRet : sRet = "<" & PoTagInfo.Item("element")
        Dim vNewArr, vArr

        '属性（attribute）の編集
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attributeが空でない場合
            Set vNewArr = new_Arr()
            Set vArr = PoTagInfo.Item("attribute").slice(0,vbNullString)
            Do While vArr.length>0
                vNewArr.push func_CmHtmlGenEditAttribute(vArr.shift)
            Loop
            sRet = sRet & " " & vNewArr.join(" ")
        End If
        
        '内容（content）の編集
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            sRet = sRet & ">"
            Set vNewArr = new_Arr()
            Set vArr = PoTagInfo.Item("content").slice(0,vbNullString)
            Do While vArr.length>0
                vNewArr.push func_CmHtmlGenEditContent(vArr.shift)
            Loop
            sRet = sRet & vNewArr.join("")
            sRet = sRet & "</" & PoTagInfo.Item("element") & ">"
        Else
        'contentが空の場合
            sRet = sRet & " />"
        End If

        '生成したHTMLを返却
        func_CmHtmlGenGenerate = sRet

        Set vNewArr = Nothing
        Set vArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenEditAttribute()
    'Overview                    : 属性（attribute）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoAttr                 : 編集する属性（attribute）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenEditAttribute( _
        byRef aoAttr _
        )
        Dim sRet
        If IsEmpty(aoAttr.Item("value")) Then
            sRet = aoAttr.Item("key")
        Else
            sRet = aoAttr.Item("key") & "=" & Chr(34) & aoAttr.Item("value") & Chr(34)
        End If
        func_CmHtmlGenEditAttribute = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenEditContent()
    'Overview                    : 内容（content）の編集処理
    'Detailed Description        : 工事中
    'Argument
    '     aoCont                 : 編集する内容（content）
    'Return Value
    '     編集結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/22         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenEditContent( _
        byRef aoCont _
        )
        Dim sRet
        On Error Resume Next
        sRet = aoCont.generate()
        If Err.Number<>0 Then
            sRet = func_CmHtmlGenHtmlEncoding(aoCont)
        End If
        On Error GoTo 0
        func_CmHtmlGenEditContent = sRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlGenHtmlEncoding()
    'Overview                    : HTMLエンコードする
    'Detailed Description        : 工事中
    'Argument
    '     asTarget               : HTMLエンコードする文字列
    'Return Value
    '     エンコードした文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/04         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlGenHtmlEncoding( _
        byRef asTarget _
        )
        Dim vSettings : vSettings = Array( _
            Array("&", "&amp;") _
            , Array("'", "&#39;") _
            , Array("""", "&quot;") _
            , Array("<", "&lt;") _
            , Array(">", "&gt;") _
            )
        Dim sTarget : sTarget = asTarget
        Dim i
        For Each i In vSettings
            sTarget = Replace(sTarget, i(0), i(1))
        Next
        func_CmHtmlGenHtmlEncoding = sTarget
    End Function

End Class
