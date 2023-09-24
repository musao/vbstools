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
'     avArr                  : 配列に追加する要素（配列）
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArraySetData( _
    byRef avArr _
    )
    Dim oArr : Set oArr = new_clsCmArray()
    oArr.PushMulti avArr
    Set new_ArraySetData = oArr
    Set oArr = Nothing
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
    Private PoArr

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
        Set PoArr = new_Dictionary()
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
        Set PoArr = Nothing
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : Property Get Item()
    'Overview                    : 配列の指定したインデックスの要素を返す
    'Detailed Description        : func_CmArrayItem()に委譲する
    'Argument
    '     alIdx                  : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Default Property Get Item( _
        byVal alIdx _
        )
        Call sub_CM_Bind(Item, func_CmArrayItem(alIdx))
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Set Item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     alIdx                  : インデックス
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
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            Call sub_CM_BindAt(PoArr, alIdx, aoEle)
        End If
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Let Item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : 工事中
    'Argument
    '     alIdx                  : インデックス
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
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            Call sub_CM_BindAt(PoArr, alIdx, aoEle)
        End If
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
        Length = PoArr.Count
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Concat()
    'Overview                    : 引数で指定した要素を連結した配列を返す
    'Detailed Description        : 自身のインスタンスは変更しない
    'Argument
    '     avArr                  : 配列に追加する要素（配列）
    'Return Value
    '     同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Concat( _
        byRef avArr _
        )
        Dim oArr : Set oArr = new_clsCmArray()
        oArr.PushMulti func_CmArrayConvArray(True)
        oArr.PushMulti avArr
        Set Concat = oArr

        Set oArr = Nothing
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
    '     avArr                  : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function PushMulti( _
        byRef avArr _
        )
        PushMulti = func_CmArrayPushMulti(avArr)
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
    'Function/Sub Name           : Sort()
    'Overview                    : 配列の要素をソートする
    'Detailed Description        : func_CM_UtilSortHeap()に委譲する
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     aoFunc                 : 判定する関数
    'Return Value
    '     ソート後の自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Sort( _
        byRef aoFunc _
        )
        Set PoArr = func_CmArrayAddDictionary(func_CM_UtilSortHeap(func_CmArrayConvArray(True), aoFunc, True), 0)
        Set Sort = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : Splice()
    'Overview                    : 配列の要素の挿入、削除、置換を行う
    'Detailed Description        : func_CmArraySplice()に委譲する
    'Argument
    '     alStart                : 開始位置のインデックス番号、負値は最後の要素のからの位置を示す
    '                              例えば-1は最後、-2は最後から2つ目の要素を示す。
    '     alDelCnt               : 開始位置から削除する要素数
    '                              0の場合は削除しない
    '     avArr                  : 開始位置に追加する要素（配列）
    'Return Value
    '     削除した配列があれば、削除した配列の同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function Splice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Set Splice = func_CmArraySplice(alStart, alDelCnt, avArr)
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
        Unshift = func_CmArrayUnshiftMulti(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : UnshiftMulti()
    'Overview                    : 配列の先頭に要素を1つ追加する
    'Detailed Description        : func_CmArrayUnshiftMulti()に委譲する
    'Argument
    '     avArr                  : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function UnshiftMulti( _
        byRef avArr _
        )
        UnshiftMulti = func_CmArrayUnshiftMulti(avArr)
    End Function





    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayItem()
    'Overview                    : 配列の指定したインデックスの要素を返す
    'Detailed Description        : 工事中
    'Argument
    '     alIdx                  : インデックス
    'Return Value
    '     指定したインデックスの要素
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayItem( _
        ByVal alIdx _
        )
        Dim oEle : Set oEle = Nothing
        If PoArr.Count>0 Then
            Call sub_CM_Bind(oEle, PoArr.Item(alIdx))
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
        Dim lIdx, vArr, lUb, boRet
        boRet = aboFlg

        '引数の関数で判定する
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If Not aoFunc(vArr(lIdx), lIdx, vArr) = aboFlg Then
                    boRet = Not aboFlg
                    Exit For
                End If
            Next
        End If

        '判定結果を返却
        func_CmArrayEveryOrSome = boRet
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
        Dim oEle, lIdx, vArr, lUb, oRet

        '引数の関数で抽出した要素だけの配列を作成
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                Call sub_CM_Bind(oEle, vArr(lIdx))
                If aoFunc(oEle, lIdx, vArr) Then
                    Call sub_CM_Push(oRet, oEle)
                End If
            Next
        End If

        '作成した配列（ディクショナリ）で当クラスのインスタンスを生成して返却
        Set func_CmArrayFilter = new_ArraySetData(oRet)

        Set oEle = Nothing
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
        Dim oEle, lIdx, vArr, lUb, oRet
        Set oRet = Nothing

        '引数の関数で抽出できる最初の要素を検索
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                Call sub_CM_Bind(oEle, vArr(lIdx))
                If aoFunc(oEle, lIdx, vArr) Then
                    Call sub_CM_Bind(oRet, oEle)
                    Exit For
                End If
            Next
        End If

        '配列から抽出した要素を返却
        Call sub_CM_Bind(func_CmArrayFind, oRet)

        Set oEle = Nothing
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
        Dim lIdx, vArr, lUb

        '配列の全ての要素について引数の関数の処理を行う
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                Call aoFunc(vArr(lIdx), lIdx, vArr)
            Next
        End If
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
        Dim oEle, lIdx, vArr, lUb, boFlg, lStart, lEnd, lStep

        '配列の全ての要素について引数の関数の処理を行う
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart=vbNullString Then
                If aboOrder Then lStart=0 Else lStart=lUb
            Else
                lStart=alStart
            End If
            If aboOrder Then lEnd=lUb Else lEnd=0
            If aboOrder Then lStep=1 Else lStep=-1

            boFlg = False
            For lIdx=lStart To lEnd Step lStep
                Call sub_CM_Bind(oEle, vArr(lIdx))

                If IsObject(avTarget) And IsObject(oEle) Then
                    If avTarget Is oEle Then boFlg = True
                ElseIf Not IsObject(avTarget) And Not IsObject(oEle) Then
                    If VarType(avTarget) = vbString And VarType(oEle) = vbString Then
                        If Strcomp(avTarget, oEle, alCompare)=0 Then boFlg = True
                    Else
                        If avTarget = oEle Then boFlg = True
                    End If
                End If

                If boFlg Then
                    func_CmArrayIndexOf = lIdx
                    Exit For
                End If

            Next
        End If

        Set oEle = Nothing
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
        Dim lIdx, vArr, lUb, vRet

        '配列の全ての要素について引数の関数の処理を行う
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                Call sub_CM_Push(vRet, aoFunc(vArr(lIdx), lIdx, vArr))
            Next
        End If

        Call sub_CM_Bind(func_CmArrayMap, new_ArraySetData(vRet))
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
        lCount = PoArr.Count
        If lCount>0 Then
            Call sub_CM_Bind(oEle, PoArr.Item(lCount-1))
            PoArr.Remove lCount-1
        End If
        Call sub_CM_Bind(func_CmArrayPop, oEle)
        Set oEle = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayPushMulti()
    'Overview                    : 配列の末尾に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     avArr                  : 配列の末尾に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayPushMulti( _
        byRef avArr _
        )
        If func_CM_ArrayIsAvailable(avArr) Then
            Dim oEle
            For Each oEle In avArr
                Call sub_CM_BindAt(PoArr, PoArr.Count, oEle)
            Next
        End If
        func_CmArrayPushMulti = PoArr.Count
        Set oEle = Nothing
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
        Dim lIdx, vArr, lUb, oRet

        '配列の全ての要素について引数の関数の処理を行う
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(aboOrder)
            lUb = Ubound(vArr)
            
            Call sub_CM_Bind(oRet, vArr(0))
            For lIdx=1 To lUb
                Call sub_CM_Bind(oRet, aoFunc(oRet, vArr(lIdx), lIdx, vArr))
            Next
        End If

        Call sub_CM_Bind(func_CmArrayReduce, oRet)

        Set oRet = Nothing
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
        If PoArr.Count>0 Then
            Set PoArr = func_CmArrayAddDictionary(func_CmArrayConvArray(False), 0)
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
        If PoArr.Count>0 Then
            '配列から取り除いた要素を返す
            Call sub_CM_Bind(func_CmArrayShift, PoArr.Item(0))
            '作成した配列（ディクショナリ）を置換え
            Set PoArr = func_CmArrayAddDictionary(func_CmArrayConvArray(True), 1)
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
        Dim lIdx, vArr, lUb, vRet, lStart, lEnd

        '配列の一部を切り出す
        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart = PoArr.Count + alStart Else lStart = alStart
            If alEnd = vbNullString Then
                lEnd = lUb
            Else
                If alEnd<0 Then lEnd = lUb + alEnd Else lEnd = alEnd - 1
            End If
            
            For lIdx=lStart To lEnd
                Call sub_CM_Push(vRet, vArr(lIdx))
            Next
        End If

        Set func_CmArraySlice = new_ArraySetData(vRet)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySplice()
    'Overview                    : 配列の要素の挿入、削除、置換を行う
    'Detailed Description        : 工事中
    'Argument
    '     alStart                : 開始位置のインデックス番号、負値は最後の要素のからの位置を示す
    '                              例えば-1は最後、-2は最後から2つ目の要素を示す。
    '     alDelCnt               : 開始位置から削除する要素数
    '                              0の場合は削除しない
    '     avArr                  : 開始位置に追加する要素（配列）
    'Return Value
    '     削除した配列があれば、削除した配列の同クラスの別インスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/23         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySplice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Dim lIdx, vArr, lUb, vArrayAft, vRet(), lStart

        If PoArr.Count>0 Then
            vArr = func_CmArrayConvArray(True)
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart = PoArr.Count + alStart Else lStart = alStart
            
            For lIdx = 0 To lStart - 1
            '開始位置までは今の配列のまま
                Call sub_CM_Push(vArrayAft, vArr(lIdx))
            Next
            
            For lIdx = lStart To lStart + alDelCnt -1
            '開始位置から削除する要素数は戻り値の配列に移す
                Call sub_CM_Push(vRet, vArr(lIdx))
            Next
            
            If func_CM_ArrayIsAvailable(avArr) Then
            '追加する要素があれば追加する
                For lIdx = 0 To Ubound(avArr)
                '削除した要素以降は今の配列に残す
                    Call sub_CM_Push(vArrayAft, avArr(lIdx))
                Next
            End If
            
            For lIdx = lStart + alDelCnt To lUb
            '削除した要素以降は今の配列に残す
                Call sub_CM_Push(vArrayAft, vArr(lIdx))
            Next
            
            
            '配列から取り除いた要素を返す
            Call sub_CM_Bind(func_CmArraySplice, new_ArraySetData(vRet))
            '作成した配列（ディクショナリ）を置換え
            Set PoArr = func_CmArrayAddDictionary(vArrayAft, 0)
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUnshiftMulti()
    'Overview                    : 配列の先頭に要素を複数追加する
    'Detailed Description        : 工事中
    'Argument
    '     avArr                  : 配列の先頭に追加する要素（配列）
    'Return Value
    '     配列内の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUnshiftMulti( _
        byRef avArr _
        )
        Dim oArr, oEle
        Set oArr = new_Dictionary()

        If func_CM_ArrayIsAvailable(avArr) Then
        '引数の要素を先頭に追加
            Set oArr = func_CmArrayAddDictionary(avArr, 0)
        End If

        '続いて今ある要素を追加
        For Each oEle In func_CmArrayConvArray(True)
            Call sub_CM_BindAt(oArr, oArr.Count, oEle)
        Next

        '作成した配列（ディクショナリ）を置換え
        Set PoArr = oArr
        func_CmArrayUnshiftMulti = PoArr.Count

        Set oEle = Nothing
        Set oArr = Nothing
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
        Dim lIdx, vArr, vRet, lStt, lEnd, lStep

        '配列の全ての要素
        If PoArr.Count>0 Then
            vArr = PoArr.Items()
            
            If aboOrder Then
                lStt = 0 : lEnd = PoArr.Count-1 : lStep = 1
            Else
                lStt = PoArr.Count-1 : lEnd = 0 : lStep = -1
            End If

            For lIdx=lStt To lEnd Step lStep
                Call sub_CM_Push(vRet, vArr(lIdx))
            Next
        End If

        func_CmArrayConvArray = vRet
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayAddDictionary()
    'Overview                    : 引数の配列の内容を配列（ディクショナリ）に追加する
    'Detailed Description        : 工事中
    'Argument
    '     avArr                  : 配列（ディクショナリ）に追加する要素（配列）
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
        byRef avArr _
        , byVal alStart _
        )
        Dim oArr, lStart, lIdx, lUb

        lUb = Ubound(avArr)
        If alStart<0 Then lStart = lUb + alStart Else lStart = alStart
        Set oArr = new_Dictionary()

        For lIdx = alStart To lUb
            Call sub_CM_BindAt(oArr, oArr.Count, avArr(lIdx))
        Next

        '作成した配列（ディクショナリ）を返す
        Set func_CmArrayAddDictionary = oArr

        Set oArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayInspectIndex()
    'Overview                    : インデックスが有効か検査する
    'Detailed Description        : 工事中
    'Argument
    '     alIdx                  : インデックス
    'Return Value
    '     結果 True:有効なインデックス / False:無効なインデックス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayInspectIndex( _
        byVal alIdx _
        )
        func_CmArrayInspectIndex = False
        If 0 <= alIdx And alIdx < PoArr.Count Then func_CmArrayInspectIndex = True
    End Function

End Class
