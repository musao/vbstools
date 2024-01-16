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
    Private PvArr,PoBroker,PlCnt

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
        Set PoBroker = Nothing
        PlCnt = 0
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
        Set PoBroker = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Set broker()
    'Overview                    : ブローカークラスのオブジェクトを設定する
    'Detailed Description        : 工事中
    'Argument
    '     aoBroker               : ブローカークラスのインスタンス
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Set broker( _
        byRef aoBroker _
        )
        Set PoBroker = aoBroker
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get broker()
    'Overview                    : ブローカークラスのオブジェクトを返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     ブローカークラスのインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get broker()
        Set broker = PoBroker
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get count()
    'Overview                    : 配列内の要素数を返す
    'Detailed Description        : func_CmArrayLength()に委譲する
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
        count = func_CmArrayLength()
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
        cf_bind item, func_CmArrayItem(alIdx)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Set item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : sub_CmArraySetLetItem()に委譲する
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
        sub_CmArraySetLetItem alIdx, aoEle
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Let item()
    'Overview                    : 配列の指定したインデックスに要素を設定する
    'Detailed Description        : sub_CmArraySetLetItem()に委譲する
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
        sub_CmArraySetLetItem alIdx, aoEle
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get items()
    'Overview                    : 配列を返す
    'Detailed Description        : func_CmArrayCopyArray()に委譲する
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
        items = func_CmArrayCopyArray(True)
    End Property

    '***************************************************************************************************
    'Function/Sub Name           : Property Get length()
    'Overview                    : 配列内の要素数を返す
    'Detailed Description        : func_CmArrayLength()に委譲する
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
        length = func_CmArrayLength()
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
        If func_CmArrayLength()>0 Then
            oArr.pushMulti PvArr
        End If
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
        cf_bind find, func_CmArrayFind(aoFunc)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : forEach()
    'Overview                    : 配列の全ての要素について引数の関数の処理を行う
    'Detailed Description        : sub_CmArrayForEach()に委譲する
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
        sub_CmArrayForEach aoFunc
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : hasElements()
    'Overview                    : 配列が要素を含むか検査する
    'Detailed Description        : func_CmArrayHasElement()に委譲する
    '                              初期状態の配列はFalseを返す
    'Argument
    '     avArr                  : 配列
    'Return Value
    '     結果 True:要素を含む / False:要素を含まない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function hasElement( _
        byRef avArr _
        )
        hasElement = func_CmArrayHasElement(avArr)
    End Function

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
    'Function/Sub Name           : join()
    'Overview                    : 配列の各要素を連結した文字列を作成する
    'Detailed Description        : vbscriptのJoin関数と同等の機能
    'Argument
    '     asDel                  : 区切り文字
    'Return Value
    '     配列の各要素を連結した文字列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/08         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function join( _
        byVal asDel _
        )
        If func_CmArrayLength()>0 Then
            join = func_CM_UtilJoin(PvArr, asDel)
        Else
            join = ""
        End If
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
        cf_bind map, func_CmArrayMap(aoFunc)
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
        cf_bind pop, func_CmArrayPop()
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
    'Overview                    : 配列の末尾に要素を複数追加する
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
    '     avInitial              : 初期値
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
        , byRef avInitial _
        )
        cf_bind reduce, func_CmArrayReduce(aoFunc, avInitial, True)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : reduceRight()
    'Overview                    : 配列のそれぞれの要素に対して逆順に引数の関数で算出した結果を返す
    'Detailed Description        : func_CmArrayReduce()に委譲する
    'Argument
    '     aoFunc                 : 関数
    '     avInitial              : 初期値
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
        , byRef avInitial _
        )
        cf_bind reduceRight, func_CmArrayReduce(aoFunc, avInitial, False)
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
        PvArr = func_CmArrayCopyArray(False)
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
        cf_bind shift, func_CmArrayShift()
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
    'Detailed Description        : func_CmArraySort()に委譲する
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
        Set sort = func_CmArraySort(Getref("func_CM_UtilSortDefaultFunc"), aboOrder)
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sortUsing()
    'Overview                    : 指定した関数を使って配列の要素をソートする
    'Detailed Description        : func_CmArraySort()に委譲する
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
        Set sortUsing = func_CmArraySort(aoFunc, True)
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
    'Function/Sub Name           : toArray()
    'Overview                    : 配列を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/16         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toArray( _
        )
        toArray = func_CmArrayCopyArray(True)
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
    '2023/12/24         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function toString( _
        )
        If func_CmArrayLength()>0 Then
            Dim vRet, oEle
            For Each oEle In PvArr
                cf_push vRet, cf_toString(oEle)
            Next
            toString = "<" & TypeName(Me) & ">[" & func_CM_UtilJoin(vRet, ",") & "]"
            Set oEle = Nothing
        Else
            toString = "<" & TypeName(Me) & ">[]"
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : uniq()
    'Overview                    : 配列の重複を排除する
    'Detailed Description        : func_CmArrayUniq()に委譲する
    'Argument
    '     なし
    'Return Value
    '     処理後の自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function uniq( _
        )
        Set uniq = func_CmArrayUniq()
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
    'Function/Sub Name           : func_CmArrayComparison()
    'Overview                    : 比較処理
    'Detailed Description        : 必要に応じて統計情報を取得する
    'Argument
    '     avEleA                 : 比較する要素A
    '     avEleB                 : 比較する要素B
    'Return Value
    '     比較結果
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayComparison( _
        byRef aoFunc _
        , byRef avEleA _
        , byRef avEleB _
        )
        If PoBroker Is Nothing Then
            cf_bind func_CmArrayComparison, aoFunc(avEleA, avEleB)
        Else
            Dim lCnt : lCnt = func_CmArrayGetCount()
            sub_CmArrayPublish "event", Array("Comparison", lCnt, "0Start", avEleA, avEleB)
            Dim vRet : cf_bind vRet, aoFunc(avEleA, avEleB)
            cf_bind func_CmArrayComparison, vRet
            sub_CmArrayPublish "event", Array("Comparison", lCnt, "1End", vRet)
        End If
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
        If func_CmArrayLength()>0 Then
            vArr = PvArr
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
        Dim lIdx, vArr, lUb, vRet

        '引数の関数で抽出した要素だけの配列を作成
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If aoFunc(vArr(lIdx), lIdx, vArr) Then
                    cf_push vRet, vArr(lIdx)
                End If
            Next
        End If
        
        '作成した配列（ディクショナリ）で当クラスのインスタンスを生成して返却
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArrayFilter = new_ArrWith(vRet)
        Else
            Set func_CmArrayFilter = new_Arr()
        End If
        
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayFind()
    'Overview                    : 引数の関数で抽出した最初の要素を返す
    'Detailed Description        : 抽出できない場合はEmptyを返す
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
        Dim lIdx, vArr, lUb, oRet
        oRet = Empty

        '引数の関数で抽出できる最初の要素を検索
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                If aoFunc(vArr(lIdx), lIdx, vArr) Then
                    cf_bind oRet, vArr(lIdx)
                    Exit For
                End If
            Next
        End If

        '配列から抽出した要素を返却
        cf_bind func_CmArrayFind, oRet

        Set oRet = Nothing
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArrayForEach()
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
    Private Sub sub_CmArrayForEach( _
        byRef aoFunc _
        )
        Dim lIdx, vArr, lUb

        '配列の全ての要素について引数の関数の処理を行う
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                aoFunc vArr(lIdx), lIdx, vArr
            Next
            PvArr = vArr
        End If
    End Sub

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayGetCount()
    'Overview                    : 連番取得
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     連番
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayGetCount( _
        )
        PlCnt = PlCnt + 1 : func_CmArrayGetCount = PlCnt
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayHasElement()
    'Overview                    : 配列が要素を含むか検査する
    'Detailed Description        : 初期状態の配列はFalseを返す
    'Argument
    '     avArr                  : 配列
    'Return Value
    '     結果 True:要素を含む / False:要素を含まない
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/15         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayHasElement( _
        byRef avArr _
        )
        func_CmArrayHasElement = False
        If IsArray(avArr) Then
            On Error Resume Next
            Dim lUb : lUb = Ubound(avArr)
            If Err.Number=0 And lUb>=0 Then func_CmArrayHasElement = True
            On Error Goto 0
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
        Dim lIdx, vArr, lUb, lStart, lEnd, lStep

        '配列の全ての要素について引数の関数の処理を行う
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            If alStart=vbNullString Then
                If aboOrder Then lStart=0 Else lStart=lUb
            Else
                lStart=alStart
            End If
            If aboOrder Then lEnd=lUb Else lEnd=0
            If aboOrder Then lStep=1 Else lStep=-1

            For lIdx=lStart To lEnd Step lStep
                If cf_isSame(avTarget, vArr(lIdx)) Then
                    func_CmArrayIndexOf = lIdx
                    Exit For
                End If
            Next
        End If
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
        byVal alIdx _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            cf_bind func_CmArrayItem, PvArr(alIdx)
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray-func_CmArrayItem()", "インデックスが有効範囲にありません。"
        End If
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
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            For lIdx=0 To lUb
                cf_push vRet, aoFunc(vArr(lIdx), lIdx, vArr)
            Next
        End If
        
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArrayMap = new_ArrWith(vRet)
        Else
            Set func_CmArrayMap = new_Arr()
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArrayPublish()
    'Overview                    : 出版（Publish）処理
    'Detailed Description        : 工事中
    'Argument
    '     asTopic                : トピック
    '     asCont                 : 内容
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArrayPublish( _
        byVal asTopic _
        , byRef avCont _
        )
        PoBroker.Publish asTopic, avCont
    End Sub

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
        Dim oRet,lUb
        oRet = Empty
        If func_CmArrayLength()>0 Then
            lUb = Ubound(PvArr)
            cf_bind oRet, PvArr(lUb)
            Redim Preserve PvArr(lUb-1)
        End If
        cf_bind func_CmArrayPop, oRet
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
        cf_pushMulti PvArr, avArr
        func_CmArrayPushMulti = func_CmArrayLength()
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
    '     avInitial              : 初期値
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
        , byRef avInitial _
        , byVal aboOrder _
        )
        Dim lIdx, vArr, lUb, oRet
        oRet = Empty

        '配列の全ての要素について引数の関数の処理を行う
        If func_CmArrayLength()>0 Then
            If aboOrder Then vArr = PvArr Else vArr = func_CmArrayCopyArray(aboOrder)
            lUb = Ubound(vArr)
            
            If IsEmpty(avInitial) Then cf_bind oRet, vArr(0) Else cf_bind oRet, avInitial
            For lIdx=1 To lUb
                cf_bind oRet, aoFunc(oRet, vArr(lIdx), lIdx, vArr)
            Next
            
            cf_bind func_CmArrayReduce, oRet
            Set oRet = Nothing
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray-func_CmArrayReduce()", "配列の初期値がありません。"
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
        If func_CmArrayLength()>0 Then
            Dim vArr : vArr = PvArr
            '配列の先頭の要素を返す
            cf_bind func_CmArrayShift, vArr(0)
            
            '先頭の要素を取り除く
            Dim lIdx, lUb
            lUb=Ubound(vArr)
            Redim vNewArr(lUb-1)
            For lIdx=1 To lUb
                cf_bind vNewArr(lIdx-1), vArr(lIdx)
            Next
            PvArr = vNewArr
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
        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart=lUb+1 Else lStart=0
            lStart = math_max(lStart+alStart,0)
            lStart = math_min(lStart,lUb+1)
            
            if alEnd=vbNullString Then
                lEnd = lUb
            Else
                If alEnd<0 Then lEnd=lUb Else lEnd=-1
                lEnd = math_max(lEnd+alEnd,-1)
                lEnd = math_min(lEnd,lUb)
            End If
            
            For lIdx=lStart To lEnd
                cf_push vRet, vArr(lIdx)
            Next
        End If
        
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArraySlice = new_ArrWith(vRet)
        Else
            Set func_CmArraySlice = new_Arr()
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySort()
    'Overview                    : 指定した関数を使って配列の要素をソートする
    'Detailed Description        : ソート処理はfunc_CM_UtilSortHeap()に委譲する
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     aoFunc                 : 関数
    '     aboOrder               : True:昇順 / False:降順
    'Return Value
    '     ソート後の自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/14         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySort( _
        byRef aoFunc _
        , byVal aboOrder _
        )
'        sub_CmArraySortBubble aoFunc, aboOrder
'        sub_CmArraySortQuick aoFunc, aboOrder
        sub_CmArraySortMerge aoFunc, aboOrder
'        sub_CmArraySortHeap aoFunc, aboOrder
        
        Set func_CmArraySort = Me
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySortBubble()
    'Overview                    : バブルソート
    'Detailed Description        : 計算回数はO(N^2)
    '                              配列の要素がないまたは1つの場合は何もしない
    'Argument
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2024/01/06         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySortBubble( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If func_CmArrayLength()<2 Then Exit Sub
        Dim vArr : vArr = PvArr
        
        Dim lEnd, lPos
        lEnd = Ubound(vArr)
        Do While lEnd>0
            For lPos=0 To lEnd-1
                If func_CmArrayComparison(aoFunc, vArr(lPos), vArr(lPos+1))=aboFlg Then
                'lPos番目の要素と(lPos+1)番目の要素を入れ替える
                    cf_swap vArr(lPos), vArr(lPos+1)
                End If
            Next
            lEnd = lEnd-1
        Loop
        PvArr = vArr
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySortQuick()
    'Overview                    : クイックソート
    'Detailed Description        : 計算回数は平均O(N*logN)、最悪はO(N^2)
    '                              配列の要素がないまたは1つの場合は何もしない
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySortQuick( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If func_CmArrayLength()<2 Then Exit Sub
        PvArr = func_CmArraySortQuick(PvArr, aoFunc, aboFlg)
    End Sub
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySortQuick()
    'Overview                    : クイックソートの再帰処理
    'Detailed Description        : sub_CmArraySortQuick()参照
    'Argument
    '     avArr                  : 配列
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     ソート後の配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySortQuick( _
        byRef avArr _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        func_CmArraySortQuick = avArr
        If Not func_CmArrayHasElement(avArr) Then Exit Function
        If Ubound(avArr)=0 Then Exit Function
        
        '0番目の要素をピボットに決める
        Dim oPivot : cf_bind oPivot, avArr(0)
        
        'ピボットと要素を関数で判定し判定方法に合致するグループをRight、そうでないグループをLeftとする
        Dim lPos, vRight, vLeft
        For lPos=1 To Ubound(avArr)
            If func_CmArrayComparison(aoFunc, avArr(lPos), oPivot)=aboFlg Then
                cf_push vRight, avArr(lPos)
            Else
                cf_push vLeft, avArr(lPos)
            End If
        Next
        
        '上述で分けたRight、Leftのグループごとに再帰処理する
        vLeft = func_CmArraySortQuick(vLeft, aoFunc, aboFlg)
        vRight = func_CmArraySortQuick(vRight, aoFunc, aboFlg)
        
        'Leftにピボット＋Rightを結合する
        cf_push vLeft, oPivot
        If func_CmArrayHasElement(vRight) Then cf_pushMulti vLeft, vRight
        
        func_CmArraySortQuick = vLeft
        Set oPivot = Nothing
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySortMerge()
    'Overview                    : マージソート
    'Detailed Description        : 計算回数はO(N*logN)
    '                              配列の要素がないまたは1つの場合は何もしない
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySortMerge( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If func_CmArrayLength()<2 Then Exit Sub
        PvArr = func_CmArraySortMerge(PvArr, aoFunc, aboFlg)
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySortMerge()
    'Overview                    : マージソートの再帰処理
    'Detailed Description        : sub_CmArraySortMerge()参照
    '                              マージ処理はfunc_CmArraySortMergeMerge()に委譲する
    'Argument
    '     avArr                  : 配列
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     ソート後の配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySortMerge( _
        byRef avArr _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        func_CmArraySortMerge = avArr
        If Not func_CmArrayHasElement(avArr) Then Exit Function
        If Ubound(avArr)=0 Then Exit Function
        
        '2つの配列に分解する
        Dim lLength, lMedian
        lLength = Ubound(avArr) - Lbound(avArr) + 1
        lMedian = math_roundUp(lLength/2, 0)
        Dim lPos, vFirst, vSecond
        For lPos=Lbound(avArr) To lMedian-1
            cf_push vFirst, avArr(lPos)
        Next
        For lPos=lMedian To Ubound(avArr)
            cf_push vSecond, avArr(lPos)
        Next
        
        '再帰処理で配列の要素が1つになるまで分解する
        vFirst = func_CmArraySortMerge(vFirst, aoFunc, aboFlg)
        vSecond = func_CmArraySortMerge(vSecond, aoFunc, aboFlg)
        
        'マージをしながら上位に戻す
        func_CmArraySortMerge = func_CmArraySortMergeMerge(vFirst, vSecond, aoFunc, aboFlg)
        
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : func_CmArraySortMergeMerge()
    'Overview                    : マージソートのマージ処理
    'Detailed Description        : func_CmArraySortMerge()から呼び出す
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     avFirst                : マージするソート済みの配列
    '     avSecond               : マージするソート済みの配列
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     マージ済の配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/18         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArraySortMergeMerge( _
        byRef avFirst _
        , byRef avSecond _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        Dim lPosF, lPosS, lEndF, lEndS
        lPosF = Lbound(avFirst) : lPosS = Lbound(avSecond)
        lEndF = Ubound(avFirst) : lEndS = Ubound(avSecond)
        
        '双方の配列の先頭の要素同士を関数で判定して戻り値の配列に追加する
        Dim vRet
        Do While lPosF<=lEndF And lPosS<=lEndS
            If func_CmArrayComparison(aoFunc, avFirst(lPosF), avSecond(lPosS))=aboFlg Then
                cf_push vRet, avSecond(lPosS)
                lPosS = lPosS + 1
            Else
                cf_push vRet, avFirst(lPosF)
                lPosF = lPosF + 1
            End If
        Loop
        
        'それぞれ残っている方の配列の要素を追加する
        Dim lPos
        If lPosF<=lEndF Then
            For lPos=lPosF To lEndF
                cf_push vRet, avFirst(lPos)
            Next
        End If
        If lPosS<=lEndS Then
            For lPos=lPosS To lEndS
                cf_push vRet, avSecond(lPos)
            Next
        End If
        
        'マージ済の配列を返す
        func_CmArraySortMergeMerge = vRet
        
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySortHeap()
    'Overview                    : ヒープソート
    'Detailed Description        : 計算回数はO(N*logN)
    '                              配列の要素がないまたは1つの場合は何もしない
    'Argument
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     ソート後の配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySortHeap( _
        byRef aoFunc _
        , byVal aboFlg _
        )
        If func_CmArrayLength()<2 Then Exit Sub
        Dim vArr : vArr = PvArr
        
        'ヒープの作成
        Dim lLb, lUb, lSize, lParent
        lLb = Lbound(vArr) : lUb = Ubound(vArr)
        lSize = lUb - lLb + 1
        '子を持つ最下部のノードから上位に向けて順番にノード単位の処理を行う
        For lParent=lSize\2-1 To lLb Step -1
            sub_CmArraySortHeapPerNodeProc vArr, lSize, lParent, aoFunc, aboFlg
        Next
        
        'ヒープの先頭（最大/最小値）を順番に取り出す
        Do While lSize>0
            'ヒープの先頭と末尾を入れ替える
            cf_swap vArr(lLb), vArr(lSize-1)
            'ヒープサイズを１つ減らして再作成
            lSize = lSize - 1
            sub_CmArraySortHeapPerNodeProc vArr, lSize, 0, aoFunc, aboFlg
        Loop

        PvArr = vArr
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySortHeapPerNodeProc()
    'Overview                    : ヒープソートのノード単位の処理
    'Detailed Description        : sub_CmArraySortHeap()から呼び出す
    '                              引数の関数の引数は以下のとおり
    '                                currentValue :配列の要素
    '                                nextValue    :次の配列の要素
    'Argument
    '     avArr                  : 配列
    '     alSize                 : ヒープのサイズ
    '     alParent               : ノードの親の配列番号
    '     aoFunc                 : 関数
    '     aboFlg                 : 判定方法
    '                                True  :昇順（関数の結果がTrueの場合に入れ替える）
    '                                False :降順（関数の結果がFalseの場合に入れ替える）
    'Return Value
    '     ソート後の配列
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/09/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySortHeapPerNodeProc( _
        byRef avArr _
        , byVal alSize _
        , byVal alParent _
        , byRef aoFunc _
        , byVal aboFlg _
        )
        Dim lRight, lLeft, lToSwap
        lLeft = alParent*2 + 1
        lRight = lLeft + 1
        lToSwap = alParent
        
        If lRight<alSize Then
        '右側の子がある場合
            If func_CmArrayComparison(aoFunc, avArr(lRight), avArr(alParent))=aboFlg Then
            '親と右側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
                lToSwap = lRight
            End If
        End If
        
        If lLeft<alSize Then
        '左側の子がある場合
            If func_CmArrayComparison(aoFunc, avArr(lLeft), avArr(lToSwap))=aboFlg Then
            '親と右側の子の勝者と左側の子の要素を関数で判定し判定方法に合致する場合は入れ替える
                lToSwap = lLeft
            End If
        End If
        
        If lToSwap<>alParent Then
            '親と子の要素を入れ替える
            cf_swap avArr(alParent), avArr(lToSwap)
            '入れ替えた子の要素以下のノードを再処理する
            sub_CmArraySortHeapPerNodeProc avArr, alSize, lToSwap, aoFunc, aboFlg
        End If
        
    End Sub

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
        Dim lIdx, vArr, lUb, vArrayAft(), vRet(), lStart

        If func_CmArrayLength()>0 Then
            vArr = PvArr
            lUb = Ubound(vArr)
            
            If alStart<0 Then lStart=lUb+1 Else lStart=0
            lStart = math_max(lStart+alStart,0)
            lStart = math_min(lStart,lUb+1)
            
            For lIdx = 0 To lStart - 1
            '開始位置までは今の配列のまま
                cf_push vArrayAft, vArr(lIdx)
            Next
            
            For lIdx = lStart To math_min(lStart+alDelCnt-1, lUb)
            '開始位置から削除する要素数は戻り値の配列に移す
                cf_push vRet, vArr(lIdx)
            Next
        End If
        
        If func_CmArrayHasElement(avArr) Then
        '追加する要素があれば追加する
            For lIdx = 0 To Ubound(avArr)
                cf_push vArrayAft, avArr(lIdx)
            Next
        End If
        
        If func_CmArrayLength()>0 Then
            For lIdx = lStart+alDelCnt To lUb
            '削除した要素以降は今の配列に残す
                cf_push vArrayAft, vArr(lIdx) 
            Next
        End If
        
        If func_CmArrayHasElement(vArrayAft) Then
            '作成した配列に置換える
            PvArr = vArrayAft
        End If
        
        '配列から取り除いた要素を返す
        If func_CmArrayHasElement(vRet) Then
            Set func_CmArraySplice = new_ArrWith(vRet)
        Else
            Set func_CmArraySplice = new_Arr()
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayUniq()
    'Overview                    : 配列の重複を排除する
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     処理後の自身のインスタンス
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/12         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayUniq( _
        )
        '重複を排除
        Dim oEle, oDic : Set oDic = new_Dic()
        For Each oEle In PvArr
            If Not oDic.Exists(oEle) Then oDic.Add oEle, Empty
        Next
        If oDic.Count<func_CmArrayLength() Then
        '重複があった場合は新しい配列を作成
            PvArr = oDic.Keys()
        End If
        '自身のインスタンスを返す
        Set func_CmArrayUniq = Me

        Set oEle = Nothing
        Set oDic = Nothing
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
        Dim vArr, lUb, lUbAdd
        lUbAdd = 0
        If func_CmArrayHasElement(avArr) Then
        '引数の要素を先頭に追加
            vArr = avArr
            lUbAdd = Ubound(avArr)
        End If

        '続いて今ある要素を追加
        If func_CmArrayLength()>0 Then
            lUb = Ubound(PvArr)
            Redim Preserve vArr(lUbAdd + func_CmArrayLength())
            For lIdx=0 To lUb
                cf_bind vArr(lUbAdd+lIdx+1), PvArr(lIdx)
            Next
        End If

        '作成した配列に置換え
        PvArr = vArr
        func_CmArrayUnshiftMulti = func_CmArrayLength()

    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayCopyArray()
    'Overview                    : 内部で保持する配列の複製を作成する
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
    Private Function func_CmArrayCopyArray( _
        aboOrder _
        )
        Dim vArr, vRet
        If func_CmArrayLength()>0 Then
            If aboOrder Then
                vRet=PvArr
            Else
                Redim vRet(func_CmArrayLength()-1)
                Dim lIdx, lIdxR : lIdxR = 0
                For lIdx=Ubound(PvArr) To 0 Step -1
                    cf_bind vRet(lIdxR), PvArr(lIdx)
                    lIdxR = lIdxR + 1
                Next
            End If
        Else
            vRet=Array()
        End If

        func_CmArrayCopyArray = vRet
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
        If func_CmArrayLength()>0 Then
            If 0<=alIdx And alIdx<=Ubound(PvArr) Then func_CmArrayInspectIndex=True
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : func_CmArrayLength()
    'Overview                    : 配列の要素数を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     配列の要素数
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/12/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Function func_CmArrayLength( _
        )
        If func_CmArrayHasElement(PvArr) Then
            func_CmArrayLength = Ubound(PvArr)+1
        Else
            func_CmArrayLength = 0
        End If
    End Function

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmArraySetLetItem()
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
    '2023/12/21         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmArraySetLetItem( _
        byVal alIdx _
        , byRef aoEle _
        )
        If func_CmArrayInspectIndex(alIdx) Then
            cf_bind PvArr(alIdx), aoEle
        Else
            Err.Raise 9, "clsCmArray.vbs:clsCmArray-sub_CmArraySetLetItem()", "インデックスが有効範囲にありません。"
        End If
    End Sub

End Class
