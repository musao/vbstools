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
        Set PoArr = new_Dic()
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
    'Function/Sub Name           : Property Get count()
    'Overview                    : 配列内の要素数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get count()
        count = PoArr.Count
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get item()
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
    Public Default Property Get item( _
        byVal alIdx _
        )
        Call cf_bind(item, func_CmArrayItem(alIdx))
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Set item()
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
    Public Property Set item( _
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            Call cf_bindAt(PoArr, alIdx, aoEle)
        End If
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Let item()
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
    Public Property Let item( _
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            Call cf_bindAt(PoArr, alIdx, aoEle)
        End If
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get items()
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
    Public Property Get items( _
        )
        items = func_CmArrayConvArray(True)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get length()
    'Overview                    : 配列内の要素数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get length()
        length = PoArr.Count
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : concat()
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
    Public Function concat( _
        byRef avArr _
        )
        Dim oArr : Set oArr = new_Arr()
        oArr.pushMulti func_CmArrayConvArray(True)
        oArr.pushMulti avArr
        Set concat = oArr

        Set oArr = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : every()
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
    Public Function every( _
        byRef aoFunc _
        )
        every = func_CmArrayEveryOrSome(aoFunc, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : filter()
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
    Public Function filter( _
        byRef aoFunc _
        )
        Set filter = func_CmArrayFilter(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : filterVbs()
    'Overview                    : 引数で指定した条件に合致する要素だけの配列を作成する
    'Detailed Description        : vbscriptのFilter関数と同等の機能
    'Argument
    '     asTarget               : 検索する文字列
    '     aboInclude             : 検索する文字列を検索対象とするか否かの区分
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
    Public Function filterVbs( _
        byVal asTarget _
        , byVal aboInclude _
        , byVal alCompare _
        )
        Set filterVbs = new_ArrWith( Filter(func_CmArrayConvArray(True), asTarget, aboInclude, alCompare) )
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : find()
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
    Public Function find( _
        byRef aoFunc _
        )
        Call cf_bind(find, func_CmArrayFind(aoFunc))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : forEach()
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
    Public Sub forEach( _
        byRef aoFunc _
        )
        Call func_CmArrayForEach(aoFunc)
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : indexOf()
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
    Public Function indexOf( _
        byRef avTarget _
        )
        indexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : joinVbs()
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
    Public Function joinVbs( _
        byVal asDelimiter _
        )
        joinVbs = Join(func_CmArrayConvArray(True), asDelimiter)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : lastIndexOf()
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
    Public Function lastIndexOf( _
        byRef avTarget _
        )
        lastIndexOf = func_CmArrayIndexOf(avTarget, vbNullString, vbBinaryCompare, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : map()
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
    Public Function map( _
        byRef aoFunc _
        )
        Call cf_bind(map, func_CmArrayMap(aoFunc))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pop()
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
    Public Function pop( _
        )
        Call cf_bind(pop, func_CmArrayPop())
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : push()
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
    Public Function push( _
        byRef aoEle _
        )
        push = func_CmArrayPushMulti(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : pushMulti()
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
    Public Function pushMulti( _
        byRef avArr _
        )
        pushMulti = func_CmArrayPushMulti(avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduce()
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
    Public Function reduce( _
        byRef aoFunc _
        )
        Call cf_bind(reduce, func_CmArrayReduce(aoFunc, True))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduceRight()
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
    Public Function reduceRight( _
        byRef aoFunc _
        )
        Call cf_bind(reduceRight, func_CmArrayReduce(aoFunc, False))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reverse()
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
    Public Sub reverse( _
        )
        Call func_CmArrayReverse()
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : shift()
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
    Public Function shift( _
        )
        Call cf_bind(shift, func_CmArrayShift())
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : slice()
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
    Public Function slice( _
        byVal alStart _
        , byVal alEnd _
        )
        Set slice = func_CmArraySlice(alStart, alEnd)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : some()
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
    Public Function some( _
        byRef aoFunc _
        )
        some = func_CmArrayEveryOrSome(aoFunc, False)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sort()
    'Overview                    : 配列の要素をソートする
    'Detailed Description        : func_CM_UtilSortHeap()に委譲する
    'Argument
    '     aboOrder               : True:昇順 / False:降順
    'Return Value
    '     ソート後の自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/09         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function sort( _
        byVal aboOrder _
        )
        Set PoArr = func_CmArrayAddDictionary(func_CM_UtilSortHeap(func_CmArrayConvArray(True), new_Func("(c,n)=>c>n"), aboOrder), 0)
        Set sort = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sortUsing()
    'Overview                    : 指定した関数を使って配列の要素をソートする
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
    Public Function sortUsing( _
        byRef aoFunc _
        )
        Set PoArr = func_CmArrayAddDictionary(func_CM_UtilSortHeap(func_CmArrayConvArray(True), aoFunc, True), 0)
        Set sortUsing = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : splice()
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
    Public Function splice( _
        byVal alStart _
        , byVal alDelCnt _
        , byRef avArr _
        )
        Set splice = func_CmArraySplice(alStart, alDelCnt, avArr)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshift()
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
    Public Function unshift( _
        byRef aoEle _
        )
        unshift = func_CmArrayUnshiftMulti(Array(aoEle))
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : unshiftMulti()
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
    Public Function unshiftMulti( _
        byRef avArr _
        )
        unshiftMulti = func_CmArrayUnshiftMulti(avArr)
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
            Call cf_bind(oEle, PoArr.Item(alIdx))
        End If
        Call cf_bind(func_CmArrayItem, oEle)
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
                Call cf_bind(oEle, vArr(lIdx))
                If aoFunc(oEle, lIdx, vArr) Then
                    Call cf_push(oRet, oEle)
                End If
            Next
        End If

        '作成した配列（ディクショナリ）で当クラスのインスタンスを生成して返却
        Set func_CmArrayFilter = new_ArrWith(oRet)

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
                Call cf_bind(oEle, vArr(lIdx))
                If aoFunc(oEle, lIdx, vArr) Then
                    Call cf_bind(oRet, oEle)
                    Exit For
                End If
            Next
        End If

        '配列から抽出した要素を返却
        Call cf_bind(func_CmArrayFind, oRet)

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
                Call cf_bind(oEle, vArr(lIdx))

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
                Call cf_push(vRet, aoFunc(vArr(lIdx), lIdx, vArr))
            Next
        End If

        Call cf_bind(func_CmArrayMap, new_ArrWith(vRet))
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
            Call cf_bind(oEle, PoArr.Item(lCount-1))
            PoArr.Remove lCount-1
        End If
        Call cf_bind(func_CmArrayPop, oEle)
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
                Call cf_bindAt(PoArr, PoArr.Count, oEle)
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
            
            Call cf_bind(oRet, vArr(0))
            For lIdx=1 To lUb
                Call cf_bind(oRet, aoFunc(oRet, vArr(lIdx), lIdx, vArr))
            Next
        End If

        Call cf_bind(func_CmArrayReduce, oRet)

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
            Call cf_bind(func_CmArrayShift, PoArr.Item(0))
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
            
            If alStart<0 Then lStart = func_CM_MathMin(PoArr.Count + alStart, 0) Else lStart = func_CM_MathMin(alStart, lUb)
'            If alStart<0 Then lStart = PoArr.Count + alStart Else lStart = alStart
            If alEnd = vbNullString Then
                lEnd = lUb
            Else
                If alEnd<0 Then lEnd = func_CM_MathMin(lUb + alEnd, 0) Else lEnd = func_CM_MathMin(alEnd - 1, lUb)
'                If alEnd<0 Then lEnd = lUb + alEnd Else lEnd = alEnd - 1
            End If
            
            For lIdx=lStart To lEnd
                Call cf_push(vRet, vArr(lIdx))
            Next
        End If

        Set func_CmArraySlice = new_ArrWith(vRet)
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
                Call cf_push(vArrayAft, vArr(lIdx))
            Next
            
            For lIdx = lStart To lStart + alDelCnt -1
            '開始位置から削除する要素数は戻り値の配列に移す
                Call cf_push(vRet, vArr(lIdx))
            Next
            
            If func_CM_ArrayIsAvailable(avArr) Then
            '追加する要素があれば追加する
                For lIdx = 0 To Ubound(avArr)
                '削除した要素以降は今の配列に残す
                    Call cf_push(vArrayAft, avArr(lIdx))
                Next
            End If
            
            For lIdx = lStart + alDelCnt To lUb
            '削除した要素以降は今の配列に残す
                Call cf_push(vArrayAft, vArr(lIdx))
            Next
            
            
            '配列から取り除いた要素を返す
            Call cf_bind(func_CmArraySplice, new_ArrWith(vRet))
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
        Set oArr = new_Dic()

        If func_CM_ArrayIsAvailable(avArr) Then
        '引数の要素を先頭に追加
            Set oArr = func_CmArrayAddDictionary(avArr, 0)
        End If

        '続いて今ある要素を追加
        For Each oEle In func_CmArrayConvArray(True)
            Call cf_bindAt(oArr, oArr.Count, oEle)
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
                Call cf_push(vRet, vArr(lIdx))
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
        Set oArr = new_Dic()

        For lIdx = alStart To lUb
            Call cf_bindAt(oArr, oArr.Count, avArr(lIdx))
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
