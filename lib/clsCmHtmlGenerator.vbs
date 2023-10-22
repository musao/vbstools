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
        Set PoTagInfo = new_DicWith("element", Empty, "attribute", Empty, "content", Empty)
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
    'Overview                    : 今の日付時刻を取得する
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
        If IsEmpty(PoTagInfo.Item("element")) Then Exit Function

        Dim sRet : sRet = "<" & PoTagInfo.Item("element")
        If Not IsEmpty(PoTagInfo.Item("attribute")) Then
        'attributeが空でない場合
            sRet = sRet & " " & newArrWith(PoTagInfo.Item("attribute")).map(new_Func("(e,i,a)=>e.Item(""key"")&""=""&e.Item(""""value"""")")).join(" ")
        End If
        
        If Not IsEmpty(PoTagInfo.Item("content")) Then
        'contentが空でない場合
            sRet = sRet & ">"
            Dim oEle, sCont
            sCont = ""
            For Each oEle In PoTagInfo.Item("content")
            'contentごとに処理する
                If func_CM_UtilIsSame(TypeName(oEle), TypeName(Me)) Then
                    sCont = sCont & oEle.generate()
                Else
                    sCont = sCont & oEle
                End If
            Next
            sRet = sRet & sCont & "</" & PoTagInfo.Item("element") & ">"
        Else
        'contentが空の場合
            sRet = sRet & " />"
        End If

        '生成したHTMLを返却
        func_CmHtmlGenGenerate = sRet
    End Function
End Class
