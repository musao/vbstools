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
'     aoEles                 : 配列に追加する要素（配列）
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySetData( _
    byRef aoEles     _
    )
    Dim oArray : Set oArray = new_clsCmArray()
    oArray.PushMulti aoEles    
    Set new_ArraySetData = oArray
    Set oArray = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArraySplit()
'Overview                    : インスタンス生成関数
'Detailed Description        : vbscriptのSplit関数と同等の機能、同クラスのインスタンスを返す
'Argument
'     asTarget               : 部分文字列と区切り文字を含む文字列表現
'     asDelimiter            : 区切り文字
'     alCompare              : 比較方法
'                                0(vbBinaryCompare):バイナリ比較
'                                1(vbTextCompare):テキスト比較
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySplit( _
    byVal asTarget _
    , byVal asDelimiter _
    , byVal alCompare _
    )
    Set new_ArraySplit = new_ArraySetData(Split(asTarget, asDelimiter, -1, alCompare))
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
    '     alIndex                : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Item( _
        byVal alIndex _
        )
        Call sub_CM_Bind(Item, func_CmArrayItem(alIndex))
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set Item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     alIndex                : インデックス
    '     aoEle                  : 設定する要素
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set Item( _
        byVal alIndex _
        , byRef aoEle _
        )
        Call sub_CM_BindAt(PoArray, alIndex, aoEle)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Let Item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     alIndex                : インデックス
    '     aoEle                  : 設定する要素
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Let Item( _
        byVal alIndex _
        , byRef aoEle _
        )
        Call sub_CM_BindAt(PoArray, alIndex, aoEle)
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get Items()
    'Overview                    : 配列を返す
    'Detailed Description        : func_CmArrayConvArray()に委譲する
    'Argument
    '     なし
    'Return Value
    '     配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get Items( _
        )
        Items = func_CmArrayConvArray(True)
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
    'Function/Sub Name           : Concat()
    'Overview                    : 引数で指定した要素を連結した配列を返す
    'Detailed Description        : 自身のインスタンスは変更しない
    'Argument
    '     aoEles                 : 配列に追加する要素（配列）
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Concat( _
        byRef aoEles     _
        )
        Dim oArray : Set oArray = new_clsCmArray()
        oArray.PushMulti func_CmArrayConvArray(True)
        oArray.PushMulti aoEles    
        Set Concat = oArray
        
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Every()
    'Overview                    : 配列の全ての要素が引数の関数の判定を満たすか確認する
    'Detailed Description        : func_CmArrayEvery()に委譲する
    'Argument
    '     aoFunc                 : 判定する関数
    'Return Value
    '     結果 True:満たす / False:満たさない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Every( _
        byRef aoFunc _
        )
        Every = func_CmArrayEveryOrSome(aoFunc, True)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Filter()
    'Overview                    : 引数の関数で抽出した要素だけの配列を作成
    'Detailed Description        : func_CmArrayFilter()に委譲する
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Filter( _
        byRef aoFunc _
        )
        Set Filter = func_CmArrayFilter(aoFunc)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : FilterVbs()
    'Overview                    : 引数で指定した条件に合致する要素だけの配列を作成する
    'Detailed Description        : vbscriptのFilter関数と同等の機能
    'Argument
    '     asTarget               : 検索する文字列
    '     aobInclude             : 検索する文字列を検索対象とするか否かの区分
    '                                True :検索する文字列を検索対象とする
    '                                False:検索する文字列以外を検索対象とする
    '     alCompare              : 比較方法
    '                                0(vbBinaryCompare):バイナリ比較
    '                                1(vbTextCompare):テキスト比較
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function FilterVbs( _
        byVal asTarget _
        , byVal aobInclude _
        , byVal alCompare _
        )
        Set FilterVbs = new_ArraySetData( Filter(func_CmArrayConvArray(True), asTarget, aobInclude, alCompare) )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Find()
    'Overview                    : 引数の関数で抽出した最初の要素を返す
    'Detailed Description        : func_CmArrayFind()に委譲する
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     配列から抽出した要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Find( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Find, func_CmArrayFind(aoFunc))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ForEach()
    'Overview                    : 配列の全ての要素について引数の関数の処理を行う
    'Detailed Description        : func_CmArrayForEach()に委譲する
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub ForEach( _
        byRef aoFunc _
        )
        Call func_CmArrayForEach(aoFunc)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : IndexOf()
    'Overview                    : 条件に合致する要素を正順に探し最初に見つかったインデックス番号を返す
    'Detailed Description        : func_CmArrayIndexOf()に委譲する
    'Argument
    '     avTarget               : 一致を確認する内容
    'Return Value
    '     条件に合致する要素のインデックス番号
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function IndexOf( _
        byRef avTarget _
        )
        IndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, True)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : JoinVbs()
    'Overview                    : 配列の各要素を連結した文字列を作成する
    'Detailed Description        : vbscriptのJoin関数と同等の機能
    'Argument
    '     asDelimiter            : 区切り文字
    'Return Value
    '     配列の各要素を連結した文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function JoinVbs( _
        byVal asDelimiter _
        )
        JoinVbs = Join(func_CmArrayConvArray(True), asDelimiter)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : LastIndexOf()
    'Overview                    : 条件に合致する要素を逆順に探し最初に見つかったインデックス番号を返す
    'Detailed Description        : func_CmArrayIndexOf()に委譲する
    'Argument
    '     avTarget               : 一致を確認する内容
    'Return Value
    '     条件に合致する要素のインデックス番号
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function LastIndexOf( _
        byRef avTarget _
        )
        LastIndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, False)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Map()
    'Overview                    : 配列から引数の関数で新たな配列を生成する
    'Detailed Description        : func_CmArrayMap()に委譲する
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Map( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Map, func_CmArrayMap(aoFunc))
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
    'Function/Sub Name           : Push()
    'Overview                    : 配列の末尾に要素を1つ追加する
    'Detailed Description        : func_CmArrayPushMulti()に委譲する
    'Argument
    '     aoEle                  : 配列の末尾に追加する要素
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Push( _
        byRef aoEle _
        )
        Push = func_CmArrayPushMulti(Array(aoEle))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : PushMulti()
    'Overview                    : 配列の末尾に要素を1つ追加する
    'Detailed Description        : func_CmArrayPushMulti()に委譲する
    'Argument
    '     aoEles                 : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function PushMulti( _
        byRef aoEles     _
        )
        PushMulti = func_CmArrayPushMulti(aoEles    )
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Reduce()
    'Overview                    : 配列のそれぞれの要素に対して正順に引数の関数で算出した結果を返す
    'Detailed Description        : func_CmArrayReduce()に委譲する
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     引数の関数で算出した結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Reduce( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(Reduce, func_CmArrayReduce(aoFunc, True))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : ReduceRight()
    'Overview                    : 配列のそれぞれの要素に対して逆順に引数の関数で算出した結果を返す
    'Detailed Description        : func_CmArrayReduce()に委譲する
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     引数の関数で算出した結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function ReduceRight( _
        byRef aoFunc _
        )
        Call sub_CM_Bind(ReduceRight, func_CmArrayReduce(aoFunc, False))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Reverse()
    'Overview                    : 配列の要素を逆順に並べる
    'Detailed Description        : func_CmArrayReverse()に委譲する
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Sub Reverse( _
        )
        Call func_CmArrayReverse()
    End Sub
    
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
    'Function/Sub Name           : Slice()
    'Overview                    : 配列の一部を切り出した配列を生成する
    'Detailed Description        : func_CmArraySlice()に委譲する
    'Argument
    '     alStart                : 開始位置のインデックス番号、負値は最後の要素のからの位置を示す
    '                              例えば-1は最後、-2は最後から2つ目の要素を示す。
    '     alEnd                  : 終了位置のインデックス番号、負値はalStartと同じ
    '                              切り出した配列に終了位置の要素は含まない
    '                              vbNullStringを指定した場合は切り出した配列に最後の要素を含める
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Slice( _
        byVal alStart _
        , byVal alEnd _
        )
        Set Slice = func_CmArraySlice(alStart, alEnd)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Some()
    'Overview                    : 配列のいずれか一つの要素が引数の関数の判定を満たすか確認する
    'Detailed Description        : func_CmArrayEvery()に委譲する
    'Argument
    '     aoFunc                 : 判定する関数
    'Return Value
    '     結果 True:満たす / False:満たさない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Some( _
        byRef aoFunc _
        )
        Some = func_CmArrayEveryOrSome(aoFunc, False)
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : Unshift()
    'Overview                    : 配列の先頭に要素を1つ追加する
    'Detailed Description        : func_CmArrayUnshiftMulti()に委譲する
    'Argument
    '     aoEle                  : 配列の先頭に追加する要素
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Unshift( _
        byRef aoEle _
        )
        Unshift = func_CmArrayUnshiftMulti(Array(aoElement))
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : UnshiftMulti()
    'Overview                    : 配列の先頭に要素を1つ追加する
    'Detailed Description        : func_CmArrayUnshiftMulti()に委譲する
    'Argument
    '     aoEles                 : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function UnshiftMulti( _
        byRef aoEles     _
        )
        UnshiftMulti = func_CmArrayUnshiftMulti(aoEles    )
    End Function
    
    
    
    
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : 配列の指定したインデックスの要素を返す
    'Detailed Description        : 工事中
    'Argument
    '     alIndex                : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        ByVal alIndex _
        )
        Dim oEle : Set oEle = Nothing
        If PoArray.Count>0 Then
            Call sub_CM_Bind(oEle, PoArray.Item(alIndex))
        End If
        Call sub_CM_Bind(func_CmArrayItem, oEle)
        Set oEle = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayEveryOrSome()
    'Overview                    : 配列の要素が引数の関数の判定を満たすか確認する
    'Detailed Description        : 引数の関数の引数は以下のとおり
    '                                element :配列の要素
    '                                index   :インデックス
    '                                array   :配列
    'Argument
    '     aoFunc                 : 判定する関数
    '     aboFlg                 : 判定方法
    '                                True  :配列の全ての要素が引数の関数の判定を満たす
    '                                False :配列のいずれか一つの要素が引数の関数の判定を満たす
    'Return Value
    '     結果 True:満たす / False:満たさない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayEveryOrSome( _
        byRef aoFunc _
        , byRef aboFlg _
        )
        Dim oItem, lIndex, oArray, boRet
        boRet = aboFlg
        
        '引数の関数で判定する
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If Not aoFunc(oItem, lIndex, oArray) = aboFlg Then
                    boRet = Not aboFlg
                    Exit For
                End If
            Next
        End If
        
        '判定結果を返却
        func_CmArrayEveryOrSome = boRet
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFilter()
    'Overview                    : 引数の関数で抽出した要素だけの配列を作成
    'Detailed Description        : 抽出できない場合は要素がない同クラスのインスタンスを返す
    '                              引数の関数の引数は以下のとおり
    '                                element :配列の要素
    '                                index   :インデックス
    '                                array   :配列
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFilter( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '引数の関数で抽出した要素だけの配列を作成
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If aoFunc(oItem, lIndex, oArray) Then
                    Call sub_CM_Push(oRet, oItem)
                End If
            Next
        End If
        
        '作成した配列（ディクショナリ）で当クラスのインスタンスを生成して返却
        Set func_CmArrayFilter = new_ArraySetData(oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFind()
    'Overview                    : 引数の関数で抽出した最初の要素を返す
    'Detailed Description        : 抽出できない場合はNothingを返す
    '                              引数の関数の引数は以下のとおり
    '                                element :配列の要素
    '                                index   :インデックス
    '                                array   :配列
    'Argument
    '     aoFunc                 : 抽出する関数
    'Return Value
    '     配列から抽出した要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/13         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayFind( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        Set oRet = Nothing
        
        '引数の関数で抽出できる最初の要素を検索
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                If aoFunc(oItem, lIndex, oArray) Then
                    Call sub_CM_Bind(oRet, oItem)
                    Exit For
                End If
            Next
        End If
        
        '配列から抽出した要素を返却
        Call sub_CM_Bind(func_CmArrayFind, oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
        Set oRet = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayForEach()
    'Overview                    : 配列の全ての要素について引数の関数の処理を行う
    'Detailed Description        : 引数の関数の引数は以下のとおり
    '                                element :配列の要素
    '                                index   :インデックス
    '                                array   :配列
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayForEach( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray
        
        '配列の全ての要素について引数の関数の処理を行う
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call aoFunc(oItem, lIndex, oArray)
            Next
        End If
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayIndexOf()
    'Overview                    : 条件に合致する要素を探し最初に見つかったインデックス番号を返す
    'Detailed Description        : 合致する要素がない場合は-1を返す
    'Argument
    '     avTarget               : 一致を確認する内容
    '     alStart                : 検索開始位置のインデックス番号
    '                              vbNullStringの場合はaboOrderが正順の場合は0、逆順の場合は全要素数-1
    '     alCompare              : 比較方法
    '                                0(vbBinaryCompare):バイナリ比較
    '                                1(vbTextCompare):テキスト比較
    '     aboOrder               : True：正順（順番どおり） / False：逆順
    'Return Value
    '     条件に合致する要素のインデックス番号
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayIndexOf( _
        byRef avTarget _
        , byVal alStart _
        , byVal alCompare _
        , byVal aboOrder _
        )
        func_CmArrayIndexOf = -1
        Dim oItem, lIndex, oArray, boFlg, lStart, lEnd, lStep
        
        '配列の全ての要素について引数の関数の処理を行う
        If PoArray.Count>0 Then
            
            If alStart=vbNullString Then
                If aboOrder Then lStart=0 Else lStart=PoArray.Count-1
            Else
                lStart=alStart
            End If
            If aboOrder Then lEnd=PoArray.Count-1 Else lEnd=0
            If aboOrder Then lStep=1 Else lStep=-1
            
            boFlg = False
            oArray = func_CmArrayConvArray(True)
            For lIndex=lStart To lEnd Step lStep
                Call sub_CM_Bind(oItem, oArray(lIndex))
                
                If IsObject(avTarget) And IsObject(oItem) Then
                    If avTarget Is oItem Then boFlg = True
                ElseIf Not IsObject(avTarget) And Not IsObject(oItem) Then
                    If VarType(avTarget) = vbString And VarType(oItem) = vbString Then
                        If Strcomp(avTarget, oItem, alCompare)=0 Then boFlg = True
                    Else
                        If avTarget = oItem Then boFlg = True
                    End If
                End If
                
                If boFlg Then
                    func_CmArrayIndexOf = lIndex
                    Exit For
                End If
                
            Next
        End If
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayMap()
    'Overview                    : 配列から引数の関数で生成した配列を返す
    'Detailed Description        : 引数の関数の引数は以下のとおり
    '                                element :配列の要素
    '                                index   :インデックス
    '                                array   :配列
    'Argument
    '     aoFunc                 : 関数
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayMap( _
        byRef aoFunc _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '配列の全ての要素について引数の関数の処理を行う
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(True)
            For lIndex=0 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, aoFunc(oItem, lIndex, oArray))
            Next
        End If
        
        Call sub_CM_Bind(func_CmArrayMap, new_ArraySetData(oRet))
        
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
        Dim oEle, lCount
        Set oEle = Nothing
        lCount = PoArray.Count
        If lCount>0 Then
            Call sub_CM_Bind(oEle, PoArray.Item(lCount-1))
            PoArray.Remove lCount-1
        End If
        Call sub_CM_Bind(func_CmArrayPop, oEle)
        Set oEle = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : 配列の末尾に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     aoEles                 : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef aoEles     _
        )
        If IsArray(aoEles    ) Then
            Dim oItem
            For Each oItem In aoEles    
                Call sub_CM_BindAt(PoArray, PoArray.Count, oItem)
            Next
        End If
        func_CmArrayPushMulti = PoArray.Count
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayReduce()
    'Overview                    : 配列のそれぞれの要素に対して引数の関数で算出した結果を返す
    'Detailed Description        : 引数の関数の引数は以下のとおり
    '                                previousValue :1つ前の配列の要素
    '                                currentValue  :配列の要素
    '                                index         :インデックス
    '                                array         :配列
    'Argument
    '     aoFunc                 : 関数
    '     aboOrder               : True：正順（順番どおり） / False：逆順
    'Return Value
    '     引数の関数で算出した結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayReduce( _
        byRef aoFunc _
        , byVal aboOrder _
        )
        Dim oItem, lIndex, oArray, oRet
        
        '配列の全ての要素について引数の関数の処理を行う
        If PoArray.Count>0 Then
            oArray = func_CmArrayConvArray(aboOrder)
            Call sub_CM_Bind(oRet, oArray(0))
            For lIndex=1 To PoArray.Count-1
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Bind(oRet, aoFunc(oRet, oItem, lIndex, oArray))
            Next
        End If
        
        Call sub_CM_Bind(func_CmArrayReduce, oRet)
        
        Set oRet = Nothing
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayReverse()
    'Overview                    : 配列の要素を逆順に並べる
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayReverse( _
        )
        If PoArray.Count>0 Then
            Set PoArray = func_CmArrayAddDictionary(func_CmArrayConvArray(False), 0)
        End If
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
        If PoArray.Count>0 Then
            '配列から取り除いた要素を返す
            Call sub_CM_Bind(func_CmArrayShift, PoArray.Item(0))
            '作成した配列（ディクショナリ）を置換え
            Set PoArray = func_CmArrayAddDictionary(func_CmArrayConvArray(True), 1)
        End If
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySlice()
    'Overview                    : 配列の一部を切り出した配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     alStart                : 開始位置のインデックス番号、負値は最後の要素のからの位置を示す
    '                              例えば-1は最後、-2は最後から2つ目の要素を示す。
    '     alEnd                  : 終了位置のインデックス番号、負値はalStartと同じ
    '                              切り出した配列に終了位置の要素は含まない
    '                              vbNullStringを指定した場合は切り出した配列に最後の要素を含める
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySlice( _
        byVal alStart _
        , byVal alEnd _
        )
        Dim oItem, lIndex, oArray, oRet, lStart, lEnd
        
        '配列の一部を切り出す
        If PoArray.Count>0 Then
            If alStart<0 Then lStart = PoArray.Count + alStart Else lStart = alStart
            If alEnd = vbNullString Then
                lEnd = PoArray.Count - 1
            Else
                If alEnd<0 Then lEnd = PoArray.Count + alEnd - 1 Else lEnd = alEnd - 1
            End If
            oArray = func_CmArrayConvArray(True)
            For lIndex=lStart To lEnd
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, oItem)
            Next
        End If
        
        Set func_CmArraySlice = new_ArraySetData(oRet)
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : 配列の先頭に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     aoEles                 : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef aoEles     _
        )
        Dim oArray, oItem
        Set oArray = new_Dictionary()
        
        If IsArray(aoEles    ) Then
        '引数の要素を先頭に追加
            Set oArray = func_CmArrayAddDictionary(aoEles    , 0)
        End If
        
        '続いて今ある要素を追加
        For Each oItem In func_CmArrayConvArray(True)
            Call sub_CM_BindAt(oArray, oArray.Count, oItem)
        Next
        
        '作成した配列（ディクショナリ）を置換え
        Set PoArray = oArray
        func_CmArrayUnshiftMulti = PoArray.Count
        
        Set oItem = Nothing
        Set oArray = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayConvArray()
    'Overview                    : 内部で保持する配列（ディクショナリ）をプリミティブの配列に変換する
    'Detailed Description        : 工事中
    'Argument
    '     aboOrder               : True：正順（順番どおり） / False：逆順
    'Return Value
    '     配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayConvArray( _
        aboOrder _
        )
        Dim oItem, lIndex, oArray, oRet, lStt, lEnd, lStep
        
        '配列の全ての要素
        If PoArray.Count>0 Then
            oArray = PoArray.Items()
            If aboOrder Then
                lStt = 0 : lEnd = PoArray.Count-1 : lStep = 1
            Else
                lStt = PoArray.Count-1 : lEnd = 0 : lStep = -1
            End If
            
            For lIndex=lStt To lEnd Step lStep
                Call sub_CM_Bind(oItem, oArray(lIndex))
                Call sub_CM_Push(oRet, oItem)
            Next
        End If
        
        func_CmArrayConvArray = oRet
        
        Set oItem = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayAddDictionary()
    'Overview                    : 引数の配列の内容を配列（ディクショナリ）に追加する
    'Detailed Description        : 工事中
    'Argument
    '     aoEles                 : 配列（ディクショナリ）に追加する要素（配列）
    '     alStart                : 開始位置のインデックス番号、負値は最後の要素のからの位置を示す
    '                              例えば-1は最後、-2は最後から2つ目の要素を示す。
    'Return Value
    '     配列（ディクショナリ）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayAddDictionary( _
        byRef aoEles     _
        , byVal alStart _
        )
        Dim oArray, lStart, lIndex, lUb
        
        lUb = Ubound(aoEles    )
        If alStart<0 Then lStart = lUb + alStart Else lStart = alStart
        Set oArray = new_Dictionary()
        
        For lIndex = alStart To lUb
            Call sub_CM_BindAt(oArray, oArray.Count, aoEles    (lIndex))
        Next
        
        '作成した配列（ディクショナリ）を返す
        Set func_CmArrayAddDictionary = oArray
        
        Set oArray = Nothing
    End Function
    
End Class
