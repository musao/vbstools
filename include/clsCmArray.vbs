'***************************************************************************************************
'FILENAME                    : clsCmArray.vbs
'Overview                    : 配列クラス
'Detailed Description        : javacsriptのArrayオブジェクト準拠、プリミティブの配列ではない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************

'***************************************************************************************************
'Function/Sub Name           : new_clsCmArray()
'Overview                    : インスタンス生成関数
'Detailed Description        : 生成した同クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_clsCmArray( _
    )
    Set new_clsCmArray = (New clsCmArray)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArraySetData()
'Overview                    : インスタンス生成関数
'Detailed Description        : 引数で指定した要素を含んだ同クラスのインスタンスを返す
'Argument
'     aoElements             : 配列に追加する要素（配列）
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySetData( _
    byRef aoElements _
    )
    Dim oArray : Set oArray = new_clsCmArray()
    oArray.PushMulti aoElements
    Set new_ArraySetData = oArray
    Set oArray = Nothing
End Function

Class clsCmArray
    'クラス内変数、定数
    Private PoArray
    
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
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Set PoArray = new_Dictionary()
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
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoArray = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Item()
    'Overview                    : 配列の指定したインデックスの要素を返す
    'Detailed Description        : func_CmArrayItem()に委譲する
    'Argument
    '     aIndex                 : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Item( _
        ByVal aIndex _
        )
        Call sub_CM_Bind(Item, func_CmArrayItem(aIndex))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Length()
    'Overview                    : 配列内の要素数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Length()
        Length = PoArray.Count
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Push()
    'Overview                    : 配列の末尾に要素を1つ追加する
    'Detailed Description        : func_CmArrayPushMulti()に委譲する
    'Argument
    '     aoElement              : 配列の末尾に追加する要素
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Push( _
        byRef aoElement _
        )
        Push = func_CmArrayPushMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : PushMulti()
    'Overview                    : 配列の末尾に要素を1つ追加する
    'Detailed Description        : func_CmArrayPushMulti()に委譲する
    'Argument
    '     aoElements             : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function PushMulti( _
        byRef aoElements _
        )
        PushMulti = func_CmArrayPushMulti(aoElements)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Unshift()
    'Overview                    : 配列の先頭に要素を1つ追加する
    'Detailed Description        : func_CmArrayUnshiftMulti()に委譲する
    'Argument
    '     aoElement              : 配列の先頭に追加する要素
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Unshift( _
        byRef aoElement _
        )
        Unshift = func_CmArrayUnshiftMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : UnshiftMulti()
    'Overview                    : 配列の先頭に要素を1つ追加する
    'Detailed Description        : func_CmArrayUnshiftMulti()に委譲する
    'Argument
    '     aoElements             : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function UnshiftMulti( _
        byRef aoElements _
        )
        UnshiftMulti = func_CmArrayUnshiftMulti(aoElements)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Pop()
    'Overview                    : 配列から末尾の要素を取り除く
    'Detailed Description        : func_CmArrayPop()に委譲する
    'Argument
    '     なし
    'Return Value
    '     配列から取り除いた要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Pop( _
        )
        Call sub_CM_Bind(Pop, func_CmArrayPop())
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Shift()
    'Overview                    : 配列から先頭の要素を取り除く
    'Detailed Description        : func_CmArrayShift()に委譲する
    'Argument
    '     なし
    'Return Value
    '     配列から取り除いた要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Shift( _
        )
        Call sub_CM_Bind(Shift, func_CmArrayShift())
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Filter()
    'Overview                    : 引数の関数で抽出した要素だけの配列を作成
    'Detailed Description        : func_CmArrayFilter()に委譲する
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     同クラスのインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Filter( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Filter, func_CmArrayFilter(aoFunc))
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : 配列の指定したインデックスの要素を返す
    'Detailed Description        : 工事中
    'Argument
    '     aIndex                 : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        ByVal aIndex _
        )
        Dim oElement : Set oElement = Nothing
        If PoArray.Count>0 Then
            Call sub_CM_Bind(oElement, PoArray.Item(aIndex))
        End If
        Call sub_CM_Bind(func_CmArrayItem, oElement)
        Set oElement = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : 配列の末尾に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     aoElements             : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef aoElements _
        )
        If IsArray(aoElements) Then
            Dim oItem
            For Each oItem In aoElements
                Call sub_CM_BindAt(PoArray, PoArray.Count, oItem)
            Next
        End If
        func_CmArrayPushMulti = PoArray.Count
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : 配列の先頭に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     aoElements             : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef aoElements _
        )
        Dim oArray, oItem
        
        If IsArray(aoElements) Then
            '引数の要素を先頭に追加
            Set oArray = new_Dictionary()
            For Each oItem In aoElements
                Call sub_CM_BindAt(oArray, oArray.Count, oItem)
            Next
        End If
        
        '続いて今ある要素を追加
        For Each oItem In PoArray.Items()
            Call sub_CM_BindAt(oArray, oArray.Count, oItem)
        Next
        
        '作成した配列（ディクショナリ）を置換え
        Set PoArray = oArray
        func_CmArrayUnshiftMulti = PoArray.Count
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPop()
    'Overview                    : 配列から末尾の要素を取り除く
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列から取り除いた要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPop( _
        )
        Dim oElement, lCount
        Set oElement = Nothing
        lCount = PoArray.Count
        If lCount>0 Then
            Call sub_CM_Bind(oElement, PoArray.Item(lCount-1))
            PoArray.Remove lCount-1
        End If
        Call sub_CM_Bind(func_CmArrayPop, oElement)
        Set oElement = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayShift()
    'Overview                    : 配列から先頭の要素を取り除く
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列から取り除いた要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayShift( _
        )
        Dim oElement, oArray, oItem, boFlg
        Set oElement = Nothing
        Set oArray = new_Dictionary()
        
        '先頭の要素を除いた配列を再作成
        If PoArray.Count>0 Then
            boFlg = False
            For Each oItem In PoArray.Items()
                If boFlg Then
                    Call sub_CM_BindAt(oArray, oArray.Count, oItem)
                Else
                    Call sub_CM_Bind(oElement, PoArray.Item(0))
                    boFlg = True
                End If
            Next
        End If
        
        '作成した配列（ディクショナリ）を置換え
        Set PoArray = oArray
        Call sub_CM_Bind(func_CmArrayShift, oElement)
        
        Set oElement = Nothing
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFilter()
    'Overview                    : 引数の関数で抽出した要素だけの配列を作成
    'Detailed Description        : 工事中
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     同クラスのインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFilter( _
        byRef aoFunc _
        )
        Dim oItem, oArray()
        
        '引数の関数で抽出した要素だけ抽出
        If PoArray.Count>0 Then
            For Each oItem In PoArray.Items()
                If aoFunc(oItem) Then
                    Call sub_CM_Push(oArray, oItem)
                End If
            Next
        End If
        
        '作成した配列（ディクショナリ）で当クラスのインスタンスを生成して返却
        Call sub_CM_Bind(func_CmArrayFilter, new_ArraySetData(oArray))
        
        Set oItem = Nothing
    End Function
    
End Class
