'***************************************************************************************************
'FILENAME                    : clsCmHtmlGenerator.vbs
'Overview                    : HTMLを生成する
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
        PoTagInfo.Item("element") = asElement
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
    'Overview                    : 属性を追加する
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

        '属性（attribute）の編集
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attributeが空でない場合
            Dim oFunc : Set oFunc = new_Func( _
            "function(e,i,a){If IsEmpty(e.Item('value')) Then:return e.Item('key'):Else:return e.Item('key') & '=''' & e.Item('value') & '''':End If}" _
            )
            sRet = sRet & " " & PoTagInfo.Item("attribute").map(oFunc).join(" ")
        End If
        
        '内容（content）の編集
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            sRet = sRet & ">"
            sRet = sRet & new_ArrWith(PoTagInfo.Item("content")).map(getref("func_CmHtmlEditContents")).join("")
            sRet = sRet & "</" & PoTagInfo.Item("element") & ">"
        Else
        'contentが空の場合
            sRet = sRet & " />"
        End If

        '生成したHTMLを返却
        func_CmHtmlGenGenerate = sRet
    End Function

'    '***************************************************************************************************
'    'Function/Sub Name           : func_CmHtmlEditAttributes()
'    'Overview                    : 属性（attribute）の編集
'    'Detailed Description        : 工事中
'    'Argument
'    '     aoEle                  : 配列の要素
'    '     alIdx                  : インデックス
'    '     avArr                  : 配列
'    'Return Value
'    '     生成したHTML
'    '---------------------------------------------------------------------------------------------------
'    'Histroy
'    'Date               Name                     Reason for Changes
'    '----------         ----------------------   -------------------------------------------------------
'    '2023/10/23         Y.Fujii                  First edition
'    '***************************************************************************************************
'    Private Function func_CmHtmlEditAttributes( _
'        byRef aoEle _
'        , byVal alIdx _
'        , byRef avArr _
'        )
'        If IsEmpty(aoEle.Item("value")) Then
'            func_CmHtmlEditAttributes = aoEle.Item("key")
'        Else
'            func_CmHtmlEditAttributes = aoEle.Item("key") & "=""" & aoEle.Item("value") & """"
'        End If
'    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmHtmlEditContents()
    'Overview                    : 内容（content）の編集
    'Detailed Description        : 工事中
    'Argument
    '     aoEle                  : 配列の要素
    '     alIdx                  : インデックス
    '     avArr                  : 配列
    'Return Value
    '     生成したHTML
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmHtmlEditContents( _
        byRef aoEle _
        , byVal alIdx _
        , byRef avArr _
        )
        If func_CM_UtilIsSame(TypeName(aoEle), TypeName(Me)) Then
            func_CmHtmlEditContents = aoEle.generate()
        Else
            func_CmHtmlEditContents = aoEle
        End If
    End Function

End Class
