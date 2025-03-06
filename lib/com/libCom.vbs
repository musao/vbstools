'***************************************************************************************************
'FILENAME                    : libCom.vbs
'Overview                    : 共通関数ライブラリ
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************

'###################################################################################################
'カスタム関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : cf_bind()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションのメンバーの場合は動作しない
'                              移送先が変数の場合に使用できる
'Argument
'     avTo                   : 移送先の変数
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bind( _
    byRef avTo _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set avTo = avValue Else avTo = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_bindAt()
'Overview                    : 変数間の項目移送
'Detailed Description        : 移送する値または変数がオブジェクトか否かによるVBS構文の違い（Setの有無）を吸収する
'                              移送先がコレクションの場合は当関数を使用する
'Argument
'     aoCollection           : 移送先のコレクション
'     asKey                  : 移送先のコレクションのキー
'     avValue                : 移送する値または変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_bindAt( _
    byRef aoCollection _
    , byVal asKey _
    , byRef avValue _
    )
    If IsObject(avValue) Then Set aoCollection.Item(asKey) = avValue Else aoCollection.Item(asKey) = avValue
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_isAvailableObject()
'Overview                    : オブジェクトが利用可能か判定する
'Detailed Description        : 工事中
'Argument
'     aoObj                  : オブジェクト
'Return Value
'     結果 True:利用可能 / False:利用不可
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isAvailableObject( _
    byRef aoObj _
    )
    Dim boFlg : boFlg = False
    If IsObject(aoObj) Then
        If Not aoObj Is Nothing Then boFlg = True
    End If
    cf_isAvailableObject = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isInteger()
'Overview                    : 整数かどうか検査する
'Detailed Description        : 工事中
'Argument
'     avValue                : 検査対象
'Return Value
'     結果 True:整数 / False:整数でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isInteger( _
    byRef avValue _
    )
    cf_isInteger = False
    If cf_isNumeric(avValue) Then cf_isInteger = (Fix(avValue) = cdbl(avValue))
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isNonNegativeNumber()
'Overview                    : 負の数でない（＝0または正の数）ことを検査する
'Detailed Description        : 工事中
'Argument
'     avValue                : 検査対象
'Return Value
'     結果 True:負の数でない / False:負の数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isNonNegativeNumber( _
    byRef avValue _
    )
    cf_isNonNegativeNumber = False
    If cf_isNumeric(avValue) Then cf_isNonNegativeNumber = Not (0 > avValue)
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isNumeric()
'Overview                    : 数値か検査する
'Detailed Description        : 工事中
'Argument
'     avValue                : 検査対象
'Return Value
'     結果 True:数値 / False:数値でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isNumeric( _
    byRef avValue _
    )
    If IsEmpty(avValue) Or IsNull(avValue) Or IsObject(avValue) Or IsArray(avValue) Then
    'Empty,Null,Object,Arrayの場合はFalse
        cf_isNumeric=False
        Exit Function
    End If
    If VarType(avValue)=vbInteger Or VarType(avValue)=vbLong Or VarType(avValue)=vbSingle Or VarType(avValue)=vbDouble Then
    'Integer,Long,Single,Doubleの場合はTrue
        cf_isNumeric=True
        Exit Function
    End If
    cf_isNumeric=False
    If VarType(avValue)=vbString Then
    'Stringの場合はIsNumeric関数の戻り値を返す
        cf_isNumeric=IsNumeric(avValue)
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isPositiveInteger()
'Overview                    : 正の整数かどうか検査する
'Detailed Description        : 工事中
'Argument
'     avValue                : 検査対象
'Return Value
'     結果 True:正の整数 / False:正の整数でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/09/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isPositiveInteger( _
    byRef avValue _
    )
    cf_isPositiveInteger = False
    If cf_isInteger(avValue) Then cf_isPositiveInteger = (cdbl(avValue) > 0)
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isSame()
'Overview                    : 同一か判定する
'Detailed Description        : 工事中
'Argument
'     avA                    : 比較対象
'     avB                    : 比較対象
'Return Value
'     結果 True:同一 / False:同一でない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isSame( _
    byRef avA _
    , byRef avB _
    )
    Dim boFlg : boFlg = False
    If IsObject(avA) And IsObject(avB) Then
        If avA Is avB Then boFlg = True
    ElseIf Not IsObject(avA) And Not IsObject(avB) Then
        If VarType(avA) = vbString And VarType(avB) = vbString Then
            If Strcomp(avA, avB, vbBinaryCompare)=0 Then boFlg = True
        Else
            If avA = avB Then boFlg = True
        End If
    End If
    cf_isSame = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_isValid()
'Overview                    : 有効な値（初期値でない）か判定する
'Detailed Description        : 工事中
'Argument
'     avTgt                  : 判定対象
'Return Value
'     結果 True:有効な値がある / False:有効な値がない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_isValid( _
    byRef avTgt _
    )
    Dim boFlg : boFlg = True
    If IsObject(avTgt) Then
    'オブジェクトの場合
        If avTgt Is Nothing Then boFlg = False
    ElseIf IsArray(avTgt) Then
    '配列の場合
        boFlg = new_Arr().hasElement(avTgt)
    Else
    '上記以外の場合
        If IsEmpty(avTgt) Or IsNull(avTgt) Then
            boFlg = False
        ElseIf cf_isSame(avTgt, vbNullString) Then
            boFlg = False
        End If
    End If
    cf_isValid = boFlg
End Function

'***************************************************************************************************
'Function/Sub Name           : cf_push()
'Overview                    : 配列に要素を追加する
'Detailed Description        : sub_CfPush()に委譲する
'Argument
'     avArr                  : 配列
'     avEle                  : 追加する要素
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_push( _
    byRef avArr _ 
    , byRef avEle _ 
    )
    sub_CfPush avArr, avEle
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_pushA()
'Overview                    : 引数の追加する配列の要素を配列に追加する
'Detailed Description        : sub_CfPushA()に委譲する
'Argument
'     avArr                  : 配列
'     avAdd                  : 追加する要素の配列
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'2024/05/01         Y.Fujii                  Rename cf_pushMulti() -> cf_pushA()
'***************************************************************************************************
Private Sub cf_pushA( _
    byRef avArr _ 
    , byRef avAdd _ 
    )
    sub_CfPushA avArr, avAdd
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_swap()
'Overview                    : 変数の値を入れ替える
'Detailed Description        : 移送処理はcf_bind()を使用する
'Argument
'     avA                    : 値を入れ替える変数
'     avB                    : 値を入れ替える変数
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/21         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub cf_swap( _
    byRef avA _
    , byRef avB _
    )
    Dim oTmp
    cf_bind oTmp, avA
    cf_bind avA, avB
    cf_bind avB, oTmp
    Set oTmp = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : cf_toString()
'Overview                    : 引数の内容を文字列で表示する
'Detailed Description        : func_CfToString()に委譲する
'Argument
'     avTgt                  : 対象
'Return Value
'     文字列に変換した引数の内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function cf_toString( _
    byRef avTgt _
    )
    cf_toString = func_CfToString(avTgt)
End Function


'***************************************************************************************************
'Function/Sub Name           : func_CfToString()
'Overview                    : 引数の内容を文字列で表示する
'Detailed Description        : 表示型式は以下のとおり
'                               配列、Dictionaryは要素ごとに内容を表示する、入れ子は再帰表示する
'                               　配列：[<Long>0,<String>"a",<Empty>,[value1,...],{key1=>value1,...},...]
'                               　Dictionary：{key1=>value1,key2=>[a_value1,...],key3=>{d_key1=>d_value1,...}...}
'                               上記以外 <VarType>Value形式 ※Valueはない場合あり
'Argument
'     avTgt                  : 対象
'Return Value
'     文字列に変換した引数の内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToString( _
    byRef avTgt _
    )
    If IsArray(avTgt) Then
        func_CfToString = func_CfToStringArray(avTgt)
        Exit Function
    End If
    If IsObject(avTgt) Then
        func_CfToString = func_CfToStringObject(avTgt)
        Exit Function
    End If
    Dim sRet : sRet = "<" & TypeName(avTgt) & ">" 
    If cf_isSame(TypeName(avTgt),"String") Then
        sRet = sRet & Chr(34) & Replace(avTgt,Chr(34),Chr(34)&Chr(34)) & Chr(34)
    ElseIf Not (IsEmpty(avTgt) Or IsNull(avTgt)) Then
        sRet = sRet & CStr(avTgt)
    End If
    func_CfToString = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringArray()
'Overview                    : 配列の内容を文字列で表示する
'Detailed Description        : 工事中
'Argument
'     avTgt                  : 対象
'Return Value
'     文字列に変換した引数の内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringArray( _
    byRef avTgt _
    )
    If new_Arr().hasElement(avTgt) Then
        Dim vRet, oEle
        For Each oEle In avTgt
            cf_push vRet, func_CfToString(oEle)
        Next
        func_CfToStringArray = "<Array>[" & Join(vRet, ",") & "]"
        Set oEle = Nothing
    Else
        func_CfToStringArray = "<Array>[]"
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringObject()
'Overview                    : オブジェクトの内容を文字列で表示する
'Detailed Description        : 工事中
'Argument
'     avTgt                  : 対象
'Return Value
'     文字列に変換した引数の内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringObject( _
    byRef avTgt _
    )
    If cf_isSame(TypeName(avTgt),"Dictionary") Then
        func_CfToStringObject = func_CfToStringObjectDictionary(avTgt)
        Exit Function
    End If

    On Error Resume Next
    func_CfToStringObject = avTgt.toString()
    If Err.Number=0 Then Exit Function
    On Error Goto 0

    If cf_isSame(VarType(avTgt), vbString) Then
        func_CfToStringObject = "<" & TypeName(avTgt) & ">" & avTgt
        Exit Function
    End If
    func_CfToStringObject = "<" & TypeName(avTgt) & ">"
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CfToStringObjectDictionary()
'Overview                    : ディクショナリの内容を文字列で表示する
'Detailed Description        : 工事中
'Argument
'     avTgt                  : 対象
'Return Value
'     文字列に変換した引数の内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CfToStringObjectDictionary( _
    byRef avTgt _
    )
    Const Cs_SPKEY = "__Special__"
    Dim sLabel : sLabel="Dictionary"
    If avTgt.Count>0 Then
        If avTgt.Exists(Cs_SPKEY) Then sLabel=avTgt.Item(Cs_SPKEY)
        Dim vRet, oEle
        For Each oEle In avTgt.Keys
            If Not cf_isSame(oEle,Cs_SPKEY) Then
                cf_push vRet, func_CfToString(oEle) & "=>" & func_CfToString(avTgt.Item(oEle))
            End If
        Next
        func_CfToStringObjectDictionary = "<" & sLabel & ">{" & Join(vRet, ",") & "}"
        Set oEle = Nothing
    Else
        func_CfToStringObjectDictionary = "<" & sLabel & ">{}"
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : sub_CfPush()
'Overview                    : 配列に要素を追加する
'Detailed Description        : 工事中
'Argument
'     avArr                  : 配列
'     avEle                  : 追加する要素
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CfPush( _
    byRef avArr _ 
    , byRef avEle _ 
    )
    On Error Resume Next
    Redim Preserve avArr(Ubound(avArr)+1)
    If Err.Number<>0 Then Redim avArr(0)
    On Error Goto 0
    cf_bind avArr(Ubound(avArr)), avEle
End Sub

'***************************************************************************************************
'Function/Sub Name           : sub_CfPushA()
'Overview                    : 引数の追加する配列の要素を配列に追加する
'Detailed Description        : 追加する配列（avAdd）が配列でない場合はsub_CfPush()を実行する
'Argument
'     avArr                  : 配列
'     avAdd                  : 追加する配列
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CfPushA( _
    byRef avArr _ 
    , byRef avAdd _ 
    )
    On Error Resume Next
    Dim lUbAdd,lIdx : lUbAdd = Ubound(avAdd)
    If Err.Number=0 Then
    '追加する配列（avAdd）が要素を持つ場合
        '配列（avArr）を拡張する
        Dim lUb : lUb = Ubound(avArr)
        If Err.Number=0 Then
            Redim Preserve avArr(lUb+lUbAdd+1)
        Else
        '配列（avArr）が要素を持たない場合はlUbを-1にする
            lUb = -1
            Redim avArr(lUbAdd)
        End If

        '配列（avArr）に追加する要素の配列（avAdd）を追加する
        For lIdx=0 To lUbAdd
            cf_bind avArr(lUb+1+lIdx), avAdd(lIdx)
        Next
    Elseif Not IsArray(avAdd) Then
    '追加する配列（avAdd）が要素を持たず配列でない場合
        sub_CfPush avArr, avAdd
    End If
    On Error Goto 0
End Sub

'###################################################################################################
'検査系の関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : ast_argFalse()
'Overview                    : 引数がFalseか検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avArgument             : 対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/28         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argFalse( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(False, avArgument) Then fw_throwException 8193, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNotEmpty()
'Overview                    : 引数がEmptyでないか検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avArgument             : 対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNotEmpty( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If IsEmpty(avArgument) Then fw_throwException 8194, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNotNull()
'Overview                    : 引数がNullでないか検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avArgument             : 対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNotNull( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If IsNull(avArgument) Then fw_throwException 8195, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argTrue()
'Overview                    : 引数がTrueか検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avArgument             : 対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/06/16         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argTrue( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(True, avArgument) Then fw_throwException 8196, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argsAreSame()
'Overview                    : 引数が同一か検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avA                    : 比較対象
'     avB                    : 比較対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argsAreSame( _
    byRef avA _
    , byRef avB _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not cf_isSame(avA, avB) Then fw_throwException 8197, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_argNull()
'Overview                    : 引数がNullか検査する
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     avArgument             : 対象
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_argNull( _
    byRef avArgument _
    , byVal asSource _
    , byVal asDescription _
    )
    If Not(IsNull(avArgument)) Then fw_throwException 8198, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : ast_failure()
'Overview                    : 例外を出す
'Detailed Description        : 同上
'Argument
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub ast_failure( _
    byVal asSource _
    , byVal asDescription _
    )
    fw_throwException 8199, asSource, asDescription
End Sub

'###################################################################################################
'フレームワーク系の関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : fw_excuteSub()
'Overview                    : 関数を実行する
'Detailed Description        : ブローカーの指定があれば実行前後に出版（Publish）処理を行う
'Argument
'     asSubName              : 実行する関数名
'     aoArg                  : 実行する関数に渡す引数
'     aoBroker               : ブローカークラスのオブジェクト
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_excuteSub( _
    byVal asSubName _
    , byRef aoArg _
    , byRef aoBroker _
    )
    Dim sSubNameForPublish : sSubNameForPublish=asSubName&"()"

    '実行前の出版（Publish） 処理
    If cf_isAvailableObject(aoBroker) Then
        aoBroker.publish topic.LOG, Array(logType.INFO ,sSubNameForPublish ,"Start")
        aoBroker.publish topic.LOG, Array(logType.DETAIL ,sSubNameForPublish ,cf_toString(aoArg))
    End If
    
    '関数の実行
    Dim oRet : Set oRet = fw_tryCatch(GetRef(asSubName), aoArg, Empty, Empty)
    
    '実行後の出版（Publish） 処理
    If cf_isAvailableObject(aoBroker) Then
        If oRet.isErr() Then
        'エラー
            aoBroker.publish topic.LOG, Array(logType.ERROR, sSubNameForPublish, cf_toString(oRet.getErr()))
        End If
        aoBroker.publish topic.LOG, Array(logType.INFO, sSubNameForPublish, "End")
        aoBroker.publish topic.LOG, Array(logType.DETAIL, sSubNameForPublish, cf_toString(aoArg))
    End If
    
    Set oRet = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_getLogPath()
'Overview                    : 実行中のスクリプトのログファイルパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのlogフォルダ以下に
'                              スクリプトファイル名＋".log"形式のファイル名で作成する
'                              fw_getPrivatePath()に委譲する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getLogPath( _
    )
    fw_getLogPath = fw_getPrivatePath("log", new_Fso().GetBaseName(WScript.ScriptName) & ".log" )
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getTextstreamForLog()
'Overview                    : log出力用のテキストストリームを作成する
'Detailed Description        : log出力先はfw_getLogPath()で取得する
'Argument
'     なし
'Return Value
'     ファイル出力バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getTextstreamForLog( _
    )
    Set fw_getTextstreamForLog = new_WriterOf(fw_getLogPath, 8, True, -1)
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getPrivatePath()
'Overview                    : 実行中のスクリプトがあるフォルダ以下のフルパスを返す
'Detailed Description        : 親フォルダ名の指定があればそのフォルダ以下のパスを返す
'                              親フォルダ名の指定がない場合は実行中のスクリプトがあるフォルダ直下のパスを返す
'Argument
'     asParentFolderName     : 親フォルダ名
'     asFileName             : ファイル名
'Return Value
'     フルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getPrivatePath( _
    byVal asParentFolderName _
    , byVal asFileName _
    )
    '実行中のスクリプトがあるフォルダのパスを取得
    Dim sParentFolderPath : sParentFolderPath = new_Fso().GetParentFolderName(WScript.ScriptFullName)
    
    'ファイルの上位ディレクトリを決める
    If Len(asParentFolderName)>0 Then
    '引数で指定したディレクトリ名がある場合
        sParentFolderPath = new_Fso().BuildPath(sParentFolderPath, asParentFolderName)
    End If

    '上位ディレクトリが存在しない場合は作成する
    fs_createFolder(sParentFolderPath)
    
    'パスを返す
    fw_getPrivatePath = new_Fso().BuildPath(sParentFolderPath, asFileName)
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_getTempPath()
'Overview                    : 一時ファイルのパスを返す
'Detailed Description        : 実行中のスクリプトがあるフォルダのtmpフォルダ以下に作成する
'                              fw_getPrivatePath()に委譲する
'Argument
'     なし
'Return Value
'     ファイルのフルパス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_getTempPath( _
    )
    fw_getTempPath = fw_getPrivatePath("tmp", new_Fso().GetTempName())
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_logger()
'Overview                    : ログ出力する
'Detailed Description        : 引数の情報にタイムスタンプを付加してファイル出力する
'Argument
'     avParams               : 配列型のパラメータリスト
'     aoWriter               : ファイル出力バッファリング処理クラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_logger( _
    byRef avParams _
    , byRef aoWriter _
    )
    Dim vIps, oEle
    For Each oEle In util_getIpAddress()
        cf_push vIps, oEle.Item("Ip").Item("V4")
    Next

    Dim oType : cf_bind oType, avParams(0)
    If cf_isSame("ReadOnlyObject", TypeName(oType)) Then avParams(0) = oType.name

    With aoWriter
        .WriteLine(new_ArrOf(Array(new_Now(), Join(vIps,","), new_Network().ComputerName)).Concat(avParams).join(vbTab))
    End With

    Set oType = Nothing
    Set oEle = Nothing
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_runShellSilently()
'Overview                    : シェルをサイレント実行する
'Detailed Description        : 同期処理、シェルの実行完了後に制御を戻す
'Argument
'     asCmd                  : 実行するコマンド
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_runShellSilently( _
    byVal asCmd _
    )
    fw_runShellSilently = False
    If 0 = new_Shell().Run(asCmd, 0, True) Then fw_runShellSilently = True
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_storeErr()
'Overview                    : Errオブジェクトの内容をオブジェクトに変換する
'Detailed Description        : 変換したオブジェクトの構成
'                              Key             Value                     例
'                              --------------  ------------------------  ---------------------------
'                              "Number"        Err.Numberの内容          11
'                              "Description"   Err.Descriptionのの内容   0 で除算しました。
'                              "Source"        Err.Sourceの内容          Microsoft VBScript 実行時エラー
'Argument
'     なし
'Return Value
'     変換したオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_storeErr( _
    )
    Dim oRet : Set oRet = new_Dic()
    '特殊キーを追加
    oRet.Add "__Special__", "Err"

    oRet.Add "Number", Err.Number
    oRet.Add "Description", Err.Description
    oRet.Add "Source", Err.Source
    Set fw_storeErr = oRet
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_storeArguments()
'Overview                    : Argumentsオブジェクトの内容をオブジェクトに変換する
'Detailed Description        : 変換したオブジェクトの構成
'                              例は引数が a /X /Hoge:Fuga, b の場合
'                              Key         Value                                        例
'                              ----------  -------------------------------------------  -------------
'                              "All"       WScript.Arguments以下のItemの内容            a /X /Hoge:Fuga, b
'                              "Named"     WScript.Arguments.Named以下のItemの内容      X: Hoge:Fuga
'                              "Unnamed"   WScript.Arguments.Unnamed以下のItemの内容    a b
'Argument
'     なし
'Return Value
'     変換したオブジェクト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_storeArguments( _
    )
    Dim oRet : Set oRet = new_Dic()
    '特殊キーを追加
    oRet.Add "__Special__", "Arguments"
    
    Dim vArr, oDic, oEle, oKey
    'All
    vArr = Array()
    For Each oEle In WScript.Arguments
        cf_push vArr, oEle
    Next
    oRet.Add "All", vArr
    
    'Named
    Set oDic = new_Dic()
    For Each oKey In WScript.Arguments.Named
        oDic.Add oKey, WScript.Arguments.Named.Item(oKey)
    Next
    oRet.Add "Named", oDic
    
    'Unnamed
    vArr = Array()
    For Each oEle In WScript.Arguments.Unnamed
        cf_push vArr, oEle
    Next
    oRet.Add "Unnamed", vArr
    
    Set fw_storeArguments = oRet
    
    Set oKey = Nothing
    Set oEle = Nothing
    Set oDic = Nothing
    Set oRet = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : fw_throwException()
'Overview                    : 例外を投げる
'Detailed Description        : 検査結果がNGの場合は例外を出す
'Argument
'     alNumber               : エラー番号
'     asSource               : ソース
'     asDescription          : 説明
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/28         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub fw_throwException( _
    byVal alNumber _
    , byVal asSource _
    , byVal asDescription _
    )
    Err.Raise alNumber, asSource, asDescription
End Sub

'***************************************************************************************************
'Function/Sub Name           : fw_tryCatch()
'Overview                    : 処理の実行とエラー発生時の処理実行
'Detailed Description        : 他の言語のtry-chatch文に準拠
'Argument
'     aoTry                  : 実行する処理（tryブロックの処理）
'     aoArgs                 : 実行する処理の引数
'     aoCatch                : エラー発生時の処理（catchブロックの処理）
'     aoFinary               : エラーの有無に依らず最後に実行する処理（finaryブロックの処理）
'Return Value
'     処理結果
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fw_tryCatch( _
    byRef aoTry _
    , byRef aoArgs _
    , byRef aoCatch _
    , byRef aoFinary _
    )
    Dim oRet, oRetF, oErr
    
    'tryブロックの処理
    On Error Resume Next
    If cf_isValid(aoArgs) Then
        cf_bind oRetF, aoTry(aoArgs)
    Else
        cf_bind oRetF, aoTry()
    End If
    Set oRet = new_Ret(oRetF)
    On Error GoTo 0

    'catchブロックの処理
    If oRet.isErr() And cf_isAvailableObject(aoCatch) Then
        If cf_isValid(aoArgs) Then
            cf_bind oRetF, aoCatch(aoArgs)
        Else
            cf_bind oRetF, aoCatch()
        End If
        if IsObject(oRetF) Then Set oRet.returnValue=oRetF Else oRet.returnValue=oRetF
    End If
    
    'finaryブロックの処理
    If cf_isAvailableObject(aoFinary) Then
        cf_bind oRetF, aoFinary(oRetF)
        if IsObject(oRetF) Then Set oRet.returnValue=oRetF Else oRet.returnValue=oRetF
    End If
    
    '結果を返却
    Set fw_tryCatch = oRet
    Set oRet = Nothing
    Set oRetF = Nothing
End Function


'###################################################################################################
'インスタンス生成関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : new_Adodb()
'Overview                    : ADOストリームオブジェクトの生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     ADOストリームオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Adodb( _
    )
    Set new_Adodb = CreateObject("ADODB.Stream")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_AdptFile()
'Overview                    : Fileオブジェクトのアダプター生成関数
'Detailed Description        : 工事中
'Argument
'     aoFile                 : ファイルのオブジェクト
'Return Value
'     生成したFileオブジェクトのアダプターのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_AdptFile( _
    byRef aoFile _
    )
    Set new_AdptFile = (New clsAdptFile).setFileObject(aoFile)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_AdptFileOf()
'Overview                    : Fileオブジェクトのアダプター生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : ファイルのパス
'Return Value
'     生成したFileオブジェクトのアダプターのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_AdptFileOf( _
    byVal asPath _
    )
    Set new_AdptFileOf = (New clsAdptFile).setFilePath(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Arr()
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
Private Function new_Arr( _
    )
    Set new_Arr = (New ArrayList)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrSplit()
'Overview                    : インスタンス生成関数
'Detailed Description        : vbscriptのSplit関数と同等の機能、同クラスのインスタンスを返す
'Argument
'     asTarget               : 部分文字列と区切り文字を含む文字列表現
'     asDelimiter            : 区切り文字
'Return Value
'     同クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/08         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ArrSplit( _
    byVal asTarget _
    , byVal asDelimiter _
    )
    Set new_ArrSplit = new_ArrOf(Split(asTarget, asDelimiter, -1, vbBinaryCompare))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ArrOf()
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
Private Function new_ArrOf( _
    byRef avArr _
    )
    Dim oArr : Set oArr = new_Arr()
    oArr.pushA avArr
    Set new_ArrOf = oArr
    Set oArr = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Broker()
'Overview                    : インスタンス生成関数
'Detailed Description        : 出版-購読型（Publish/Subscribe）クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     出版-購読型（Publish/Subscribe）クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Broker( _
    )
    Set new_Broker = (New Broker)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_BrokerOf()
'Overview                    : インスタンス生成関数
'Detailed Description        : 出版-購読型（Publish/Subscribe）クラスに指定したtopicをsubscribeして返す
'Argument
'     avParams               : 奇数（1,3,5,...）番目はtopic、偶数（2,4,6,...）番目はコールバック関数ポインタ
'                              topicだけの場合はコールバック関数ポインタをsubscribeしない
'Return Value
'     出版-購読型（Publish/Subscribe）クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/06         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_BrokerOf( _
    byVal avParams _
    )
    Dim i,vTmp,lCnt,oBroker
    Set oBroker = New Broker
    lCnt = 0
    For Each i In avParams
        lCnt = lCnt + 1
        cf_push vTmp, i
        If lCnt Mod 2 = 0 Then oBroker.subscribe vTmp(lCnt-2), vTmp(lCnt-1)
    Next
    Set new_BrokerOf = oBroker
End Function

'***************************************************************************************************
'Function/Sub Name           : new_CalAt()
'Overview                    : インスタンス生成関数
'Detailed Description        : 指定した日付時刻で生成した日付クラスのインスタンスを返す
'Argument
'     avDateTime             : 設定する日付時刻
'Return Value
'     日付クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_CalAt( _
    ByVal avDateTime _
    )
    Set new_CalAt = (New Calendar).of(avDateTime)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Char()
'Overview                    : インスタンス生成関数
'Detailed Description        : 文字種類管理クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     文字種類管理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/31         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Char( _
    )
    Set new_Char = (New CharacterType)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_CssOf()
'Overview                    : インスタンス生成関数
'Detailed Description        : CSS生成クラスのインスタンスを返す
'Argument
'     asSelector             : セレクタ
'Return Value
'     CSS生成クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_CssOf( _
    byVal asSelector _
    )
    Dim oCss : Set oCss = New CssGenerator
    oCss.selector = asSelector
    Set new_CssOf = oCss
    Set oCss = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Dic()
'Overview                    : Dictionaryオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Dic( _
    )
    Set new_Dic = CreateObject("Scripting.Dictionary")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DicOf()
'Overview                    : Dictionaryオブジェクトを生成し初期値を設定する
'Detailed Description        : 工事中
'Argument
'     avParams               : 初期値奇数（1,3,5,...）はKey、偶数（2,4,6,...）はValue
'                              Keyだけの場合は値にEmptyを設定する。
'Return Value
'     生成したDictionaryオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DicOf( _
    byVal avParams _
    )
    Dim oDict, vItem, vKey, boIsKey
    
    boIsKey = True
    Set oDict = new_Dic()
    
    For Each vItem In avParams
        If boIsKey Then
            cf_bind vKey, vItem
            cf_bindAt oDict, vKey, Empty
        Else
            cf_bindAt oDict, vKey, vItem
        End If
        boIsKey = Not boIsKey
    Next
    
    Set new_DicOf = oDict
    Set oDict = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_DriveOf()
'Overview                    : Driveオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asDriveName            : ドライブ名
'Return Value
'     生成したDriveオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_DriveOf( _
    byVal asDriveName _
    )
    Set new_DriveOf = new_Fso().GetDrive(asDriveName)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Enum()
'Overview                    : Enum生成関数
'Detailed Description        : Enumのインスタンスを生成する
'Argument
'     asName                 : Enum名
'     aoDef                  : Enumの定義
'                              定義はkeyが定義名、valueが値のDictionary
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/05/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub new_Enum( _
    byVal asName _
    , byRef aoDef _
    )
    'クラス名（仮名）作成
    Dim sClassName : sClassName = "clsTmp_" & new_Fso().GetBaseName(new_Fso().GetTempName())
'    With new_Char()
'        Dim vCharList : vCharList = .charList(.typeHalfWidthAlphabetUppercase + .typeHalfWidthNumbers)
'    End With
'    cf_push vCharList, "_"
'    Dim sClassName : sClassName = "clsTmp_" & util_randStr(vCharList, 10)

    'クラス定義のソースコード作成
    Dim sThisName : sThisName = asName
    Dim vCode,i
    cf_push vCode, "Class " & sClassName
    cf_push vCode, "Private " & Join(aoDef.Keys,"_,")&"_"
    cf_push vCode, "Private PoLists"
    cf_push vCode, "Private Sub Class_Initialize()"
    cf_push vCode, "    Set PoLists = CreateObject('Scripting.Dictionary')"
    For Each i in aoDef.Keys
        cf_push vCode, "    Set " & i & "_ = (new ReadOnlyObject).of(Me, '" & i & "', " & aoDef.Item(i) & ")"
        cf_push vCode, "    cf_bindAt PoLists, '" & i & "', " & i
    Next
    cf_push vCode, "End Sub"
    For Each i in aoDef.Keys
        cf_push vCode, "Public Property Get " & i & "()"
        cf_push vCode, "    cf_bind " & i & ", " & i & "_"
        cf_push vCode, "End Property"
    Next
    cf_push vCode, "Public Property Get toString()"
    cf_push vCode, "    Dim i,ar"
    cf_push vCode, "    For Each i In PoLists.Items"
    cf_push vCode, "        cf_push ar, i.toString()"
    cf_push vCode, "    Next"
    cf_push vCode, "    toString = '<'&TypeName(Me)&'>(" & sThisName & "){'&Join(ar,',')&'}'"
    cf_push vCode, "End Property"
    cf_push vCode, "Public Function values()"
    cf_push vCode, "    values = PoLists.Items"
    cf_push vCode, "End Function"
    cf_push vCode, "Public Function valueOf(n)"
    cf_push vCode, "    ast_argTrue PoLists.Exists(n), TypeName(Me)&'(" & sThisName & ")+valueOf()', 'There is no element with the specified name'"
    cf_push vCode, "    Set valueOf = PoLists.Item(n)"
    cf_push vCode, "End Function"
    cf_push vCode, "End Class"
    'インスタンス生成のソースコード作成
    cf_push vCode, "Private " & sThisName
    cf_push vCode, "Set " & sThisName & " = new " & sClassName
    '実行
    ExecuteGlobal Replace(Join(vCode,":"), "'", """")

End Sub

'***************************************************************************************************
'Function/Sub Name           : new_FileOf()
'Overview                    : Fileオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : パス
'Return Value
'     生成したFileオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FileOf( _
    byVal asPath _
    )
    Set new_FileOf = new_Fso().GetFile(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_FolderItem2Of()
'Overview                    : FolderItem2オブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : パス
'Return Value
'     生成したFolderItem2オブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/02/22         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FolderItem2Of( _
    byVal asPath _
    )
    With new_Fso()
        Set new_FolderItem2Of = new_ShellApp().Namespace(.GetParentFolderName(asPath)).Items().Item(.GetFileName(asPath))
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : new_FolderOf()
'Overview                    : Folderオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : パス
'Return Value
'     生成したFolderオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/23         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_FolderOf( _
    byVal asPath _
    )
    Set new_FolderOf = new_Fso().GetFolder(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Fso()
'Overview                    : FileSystemObjectオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したFileSystemObjectオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/13         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Fso( _
    )
    Set new_Fso = CreateObject("Scripting.FileSystemObject")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Func()
'Overview                    : 関数のインスタンスを生成する
'Detailed Description        : javascriptの無名関数に準拠（vbscriptの仕様上仮の名前はつける）
'Argument
'     asSoruceCode           : 生成する関数のソースコード
'                              以下のいずれかの様式とし、function（subではない）を生成する
'                              1.通常
'                               function (①) {②}
'                                ①引数をカンマ区切りで指定する
'                                ②vbscriptの構文に準拠する、戻り値は"return hoge"と表記する
'                                  "return"句がない場合は戻り値はなしとする
'                              2.Arrow関数
'                               ① => ②
'                                ①引数をカンマ区切りで指定する、複数の場合は()で囲む
'                                ②単一行の場合はそのまま戻り値とする、複数行の場合は1.通常の②と同じ
'Return Value
'     生成した関数のインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Func( _
    byVal asSoruceCode _
    )
    '生成する関数のソースコードの改行を:に変換
    Dim sSoruceCode
    sSoruceCode = Replace(asSoruceCode, vbCrLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbLf, ":")
    sSoruceCode = Replace(sSoruceCode, vbCr, ":")
    '生成する関数のソースコードの'（シングルクォーテーション）を"（ダブルクォーテーション）に変換
    sSoruceCode = Replace(sSoruceCode, "'", """")
    
    '関数名（仮名）を作る
    Dim sFuncName : sFuncName = "anonymous_" & new_Fso().GetBaseName(new_Fso().GetTempName())
'    With new_Char()
'        Dim vCharList : vCharList = .charList(.typeHalfWidthAlphabetUppercase + .typeHalfWidthNumbers)
'    End With
'    cf_push vCharList, "_"
'    Dim sFuncName : sFuncName = "anonymous_" & util_randStr(vCharList, 10)
    
    Dim sPattern, oRegExp, sArgStr, sProcStr
    '生成する関数のソースコードの様式が「1.通常」の場合
    sPattern = "function\s*\((.*)\)\s*{(.*)}"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'return句があれば関数名で書き換える
        sProcStr = func_NewRewriteReturnPhrase(sFuncName, False, func_NewAnalyze(sProcStr) )
        
        '関数の生成
        Set new_Func = func_NewGenerate(sFuncName, sArgStr, sProcStr)
        Set oRegExp = Nothing
        Exit Function
    End If
    
    '生成する関数のソースコードの様式が「2.Arrow関数」の場合
    sPattern = "(.*)\s*=>\s*(.*)\s*"
    Set oRegExp = new_Re(sPattern, "igm")
    If oRegExp.Test(sSoruceCode) Then
        sArgStr = oRegExp.Replace(sSoruceCode, "$1")
        sProcStr = oRegExp.Replace(sSoruceCode, "$2")
        
        'それぞれ前後の括弧があれば除去
        sPattern = "\(\s?(.*)\s?\)"
        sArgStr = new_Re(sPattern, "igm").Replace(sArgStr, "$1")
        sPattern = "{\s?(.*)\s?}"
        sProcStr = new_Re(sPattern, "igm").Replace(sProcStr, "$1")
        
        'return句があれば関数名で書き換える
        sProcStr = func_NewRewriteReturnPhrase(sFuncName, True, func_NewAnalyze(sProcStr) )
        
        '関数の生成
        Set new_Func = func_NewGenerate(sFuncName, sArgStr, sProcStr)
    End If
    Set oRegExp = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_HtmlOf()
'Overview                    : インスタンス生成関数
'Detailed Description        : HTML生成クラスのインスタンスを返す
'Argument
'     asElement              : 要素
'Return Value
'     HTML生成クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_HtmlOf( _
    byVal asElement _
    )
    Dim oHtml : Set oHtml = New HtmlGenerator
    oHtml.element = asElement
    Set new_HtmlOf = oHtml
    Set oHtml = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Network()
'Overview                    : WScript.Networkオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したWScript.Networkオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Network( _
    )
    Set new_Network = CreateObject("WScript.Network")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Now()
'Overview                    : インスタンス生成関数
'Detailed Description        : 今の日付時刻で生成した日付クラスのインスタンスを返す
'Argument
'     なし
'Return Value
'     日付クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Now( _
    )
    Set new_Now = (New Calendar).ofNow()
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Re()
'Overview                    : 正規表現オブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     asPattern              : 正規表現のパターン
'     asOptions              : この引数内にある文字の有無で正規表現の以下のプロパティをTrueにする
'                                "i":大文字と小文字を区別する（.IgnoreCase = True）
'                                "g"文字列全体を検索する（.Global = True）
'                                "m"文字列を複数行として扱う（.Multiline = True）
'Return Value
'     生成した正規表現オブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Re( _
    byVal asPattern _
    , byVal asOptions _
    )
    Dim oRe, sOpts
    
    Set oRe = New RegExp
    oRe.Pattern = asPattern
    
    sOpts = LCase(asOptions)
    If InStr(sOpts, "i") > 0 Then oRe.IgnoreCase = True
    If InStr(sOpts, "g") > 0 Then oRe.Global = True
    If InStr(sOpts, "m") > 0 Then oRe.Multiline = True
    
    Set new_Re = oRe
    Set oRe = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Reader()
'Overview                    : ファイル読込バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     aoTextStream           : テキストストリームオブジェクト
'Return Value
'     生成したファイル読込バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/02         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Reader( _
    byRef aoTextStream _
    )
    Set new_Reader = (New BufferedReader).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ReaderOf()
'Overview                    : ファイル読込バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : 読み込むファイルのパス
'Return Value
'     生成したファイル読込バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ReaderOf( _
    byVal asPath _
    )
    Set new_ReaderOf = (New BufferedReader).setTextStream(new_Ts(asPath, 1, False, -2))
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Ret()
'Overview                    : 戻り値クラスオブジェクトの生成関数
'Detailed Description        : 工事中
'Argument
'     avRet                  : 戻り値
'Return Value
'     生成した戻り値クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Ret( _
    byRef avRet _
    )
    Set new_Ret = (New ReturnValue).setValue(avRet)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_RetByState()
'Overview                    : 戻り値クラスオブジェクトの生成関数
'Detailed Description        : 工事中
'Argument
'     avNormal               : 正常の場合の戻り値
'     avAbnormal             : 異常の場合の戻り値
'Return Value
'     生成した戻り値クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/04/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_RetByState( _
    byRef avNormal _
    , byRef avAbnormal _
    )
    Set new_RetByState = (New ReturnValue).setValueByState(avNormal,avAbnormal)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Shell()
'Overview                    : Wscript.Shellオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したWscript.Shellオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Shell( _
    )
    Set new_Shell = CreateObject("Wscript.Shell")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_ShellApp()
'Overview                    : Shell.Applicationオブジェクト生成関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     生成したShell.Applicationオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_ShellApp( _
    )
    Set new_ShellApp = CreateObject("Shell.Application")
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Ts()
'Overview                    : TextStreamオブジェクト生成関数
'Detailed Description        : FileSystemObjectのOpenTextFile()と同等
'Argument
'     asPath                 : パス
'     alIomode               : 入力/出力モード 1:ForReading,2:ForWriting,8:ForAppending
'     aboCreate              : asPathが存在しない場合 True:新しいファイルを作成する、False:作成しない
'     alFileFormat           : ファイルの形式 -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     生成したTextStreamオブジェクトのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Ts( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_Ts = new_Fso().OpenTextFile(asPath, alIomode, aboCreate, alFileFormat)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_Writer()
'Overview                    : ファイル出力バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     aoTextStream           : テキストストリームオブジェクト
'Return Value
'     生成したファイル出力バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_Writer( _
    byRef aoTextStream _
    )
    Set new_Writer = (New BufferedWriter).setTextStream(aoTextStream)
End Function

'***************************************************************************************************
'Function/Sub Name           : new_WriterOf()
'Overview                    : ファイル出力バッファリング処理クラスのインスタンス生成関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : 書き込むファイルのパス
'     alIomode               : 出力モード 2:ForWriting,8:ForAppending
'     aboCreate              : asPathが存在しない場合 True:新しいファイルを作成する、False:作成しない
'     alFileFormat           : ファイルの形式 -2:TristateUseDefault,-1:TristateTrue,0:TristateFalse
'Return Value
'     生成したファイル出力バッファリング処理クラスのインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/09         Y.Fujii                  First edition
'***************************************************************************************************
Private Function new_WriterOf( _
    byVal asPath _
    , byVal alIomode _
    , byVal aboCreate _
    , byVal alFileFormat _
    )
    Set new_WriterOf = (New BufferedWriter).setTextStream(new_Ts(asPath, alIomode, aboCreate, alFileFormat))
End Function

'---------------------------------------------------------------------------------------------------

'***************************************************************************************************
'Function/Sub Name           : func_NewAnalyze()
'Overview                    : ソースコードを解釈する
'Detailed Description        : new_Func()から使用する
'                              _（アンダーライン）は行を結合する
'Argument
'     asCode                 : ソースコード
'Return Value
'     ソースコードを行ごとに分解した配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewAnalyze( _
    byVal asCode _
    )
    Dim sRow, sPtn, oCode, sTemp
    Set oCode = new_Dic()
    sTemp= ""
    For Each sRow In Split(asCode, ":", -1, vbBinaryCompare)
        If Len(Trim(sRow))>0 Then
            sPtn = "^(.*)\s_\s*$"
            If new_Re(sPtn, "ig").Test(sRow) Then
                sTemp = sTemp & Trim(new_Re(sPtn, "ig").Replace(sRow, "$1"))
            Else
                sRow = sTemp & " " & Trim(sRow)
                sTemp = ""
                oCode.Add oCode.Count, Trim(sRow)
            End If
        End If
    Next
    
    func_NewAnalyze = oCode.Items()
    Set oCode = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_NewGenerate()
'Overview                    : 引数の情報で関数のインスタンスを生成する
'Detailed Description        : new_Func()から使用する
'Argument
'     asFuncName             : 関数名
'     asArgStr               : ソースの引数部分のソースコード
'     asProcStr              : ソースの処理内容部分のソースコード
'Return Value
'     生成した関数のインスタンス
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewGenerate( _
    byVal asFuncName _
    , byVal asArgStr _
    , byVal asProcStr _
    )
    Dim sCode
    'ソースコード作成
    sCode = _
        "Private Function " & asFuncName & "(" & asArgStr & ")" & vbNewLine _
        & asProcStr & vbNewLine _
        & "End Function"
    
'inputbox "","",sCode
    '関数の生成
    ExecuteGlobal sCode
    Set func_NewGenerate = Getref(asFuncName)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_NewRewriteReturnPhrase()
'Overview                    : return句を書き換える
'Detailed Description        : new_Func()から使用する
'                              Arrow関数で1行の場合はその行全体をreturnする
'Argument
'     asFuncName             : 関数名
'     aboArrowFlg            : Arrow関数か否かのフラグ
'     avCode                 : ソースコードを行ごとに分解した配列
'Return Value
'     書き換えたソースの処理内容部分のソースコード
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_NewRewriteReturnPhrase( _
    byVal asFuncName _
    , byVal aboArrowFlg _
    , byRef avCode _
    )
    Dim sPtnRet : sPtnRet = "^(.*\s+)?return\s+(.*)\s{0,}$"
    
    If Ubound(avCode)=0 And aboArrowFlg=True Then
    'Arrow関数で1行の場合
        Dim sCode : sCode = avCode(0)
        If new_Re(sPtnRet, "ig").Test(sCode) Then
        'return句がある場合
            func_NewRewriteReturnPhrase = new_Re(sPtnRet, "ig").Replace(sCode, "$1 cf_bind " & asFuncName & ", ($2)")
        Else
        'return句がない場合
            func_NewRewriteReturnPhrase = "cf_bind " & asFuncName & ", (" & sCode & ")"
        End If
        Exit Function
    End If
    
    Dim lCnt, sPtn, sRow
    For lCnt=0 To Ubound(avCode)
        sRow = avCode(lCnt)
        If new_Re(sPtnRet, "ig").Test(sRow) Then
            avCode(lCnt) = new_Re(sPtnRet, "ig").Replace(sRow, "$1 cf_bind " & asFuncName & ", ($2)")
        End If
    Next
    
    func_NewRewriteReturnPhrase = Join(avCode, ":")
    
End Function

'###################################################################################################
'数学系の関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : math_min()
'Overview                    : 最小値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が小さい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_min( _
    byVal al1 _ 
    , byVal al2 _
    )
    cf_bind math_min, math_minA(Array(al1, al2))
'    Dim lRet
'    If al1 < al2 Then lRet = al1 Else lRet = al2
'    math_min = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_minA()
'Overview                    : 最小値を求める
'Detailed Description        : 工事中
'Argument
'     avNums                 : 数値
'Return Value
'     avNumsの最小値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_minA( _
    byRef avNums _
    )
    cf_bind math_minA, avNums
    If new_Arr().hasElement(avNums) Then cf_bind math_minA, new_ArrOf(avNums).sort(True)(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_max()
'Overview                    : 最大値を求める
'Detailed Description        : 工事中
'Argument
'     al1                    : 数値1
'     al2                    : 数値2
'Return Value
'     al1とal2の値が大きい方
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_max( _
    byVal al1 _ 
    , byVal al2 _
    )
    cf_bind math_max, math_maxA(Array(al1, al2))
'    Dim lRet
'    If al1 > al2 Then lRet = al1 Else lRet = al2
'    math_max = lRet
End Function

'***************************************************************************************************
'Function/Sub Name           : math_maxA()
'Overview                    : 最大値を求める
'Detailed Description        : 工事中
'Argument
'     avNums                 : 数値
'Return Value
'     avNumsの最大値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/01         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_maxA( _
    byRef avNums _
    )
    cf_bind math_maxA, avNums
    If new_Arr().hasElement(avNums) Then cf_bind math_maxA, new_ArrOf(avNums).sort(False)(0)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundUp()
'Overview                    : 切り上げする
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、切り上げする端数の位置を小数の位で表す
'Return Value
'     切り上げした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundUp( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundUp = func_MathRound(adbNum, alPlace, 9, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_round()
'Overview                    : 四捨五入する
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、四捨五入する端数の位置を小数の位で表す
'Return Value
'     四捨五入した値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_round( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_round = func_MathRound(adbNum, alPlace, 5, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_roundDown()
'Overview                    : 切り捨てする
'Detailed Description        : func_MathRound()に委譲する
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、切り捨てする端数の位置を小数の位で表す
'Return Value
'     切り捨てした値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_roundDown( _
    byVal adbNum _ 
    , byVal alPlace _
    )
    math_roundDown = func_MathRound(adbNum, alPlace, 0, True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_rand()
'Overview                    : 乱数を生成する
'Detailed Description        : 工事中
'Argument
'     adbMin                 : 生成する乱数の最小値
'     adbMax                 : 生成する乱数の最大値
'     alPlace                : 小数の位、切り上げする端数の位置を小数の位で表す
'Return Value
'     生成した乱数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_rand( _
    byVal adbMin _
    , byVal adbMax _
    , byVal alPlace _
    )
    Randomize
    math_rand = adbMin + Fix( ((adbMax-adbMin)*(10^alPlace)+1)*Rnd )*10^(-1*alPlace)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_log2()
'Overview                    : 2が底の対数
'Detailed Description        : 工事中
'Argument
'     adbAntilogarithm       : 真数
'Return Value
'     冪指数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_log2( _
    byVal adbAntilogarithm _
    )
    math_log2 = func_MathLog(2, adbAntilogarithm)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_tranc()
'Overview                    : 整数部を返す
'Detailed Description        : ゼロ方向に丸めた整数部を返す
'                               10.8  -> 10
'                               -10.8 -> -10
'Argument
'     adbNum                 : 数値
'Return Value
'     整数部
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_tranc( _
    byVal adbNum _ 
    )
    math_tranc = func_MathRound(adbNum,0,0,True)
End Function

'***************************************************************************************************
'Function/Sub Name           : math_fractional()
'Overview                    : 小数部を返す
'Detailed Description        : ゼロ方向に丸めた小数部を返す
'                               10.8  -> 0.8
'                               -10.8 -> -0.8
'Argument
'     adbNum                 : 数値
'Return Value
'     小数部
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/02/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function math_fractional( _
    byVal adbNum _ 
    )
    math_fractional = adbNum-func_MathRound(adbNum,0,0,True)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_MathRound()
'Overview                    : 数値を丸める
'Detailed Description        : 端数処理の方法に従って数値を丸める
'                              引数のalPlaceは丸めたい小数の位を指定する、例えば第一位を場合は0を指定する
'                              小数第二位を四捨五入する場合、alPlaceに1、alThresholdに5を指定する
'                              一の位、十の位、・・・の場合は-1,-2,…のように負値を指定する
'                                例）１８２．７３２
'                                　　↑　　↑　↑　 ↑
'                                   -3  -1　0　 2
'Argument
'     adbNum                 : 数値
'     alPlace                : 小数の位、処理する端数の位置を小数の位で表す
'     alThreshold            : 閾値
'                               0：切り捨て
'                               5：四捨五入
'                               9：切り上げ
'     aboMode                : 端数処理の方法
'                               True  ：符号を無視して絶対値を丸める（正負で丸める方向が異なる）
'                               False ：正数の場合と増減を同じ向きに丸める
'                              https://ja.wikipedia.org/wiki/%E7%AB%AF%E6%95%B0%E5%87%A6%E7%90%86
'Return Value
'     丸めた値
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_MathRound( _
    byVal adbNum _ 
    , byVal alPlace _
    , byVal alThreshold _
    , byVal aboMode _
    )
    Dim lThreshold : lThreshold = alThreshold
    If adbNum<0 Then lThreshold = -1*lThreshold

    Dim dbTemp
    dbTemp = Cstr((adbNum+lThreshold*10^(-1*(alPlace+1))) * 10^(alPlace))

    If aboMode Then
        func_MathRound = Cdbl( Cstr( Fix(dbTemp) * 10^(-1*alPlace) ) )
    Else
        func_MathRound = Cdbl( Cstr( Int(dbTemp) * 10^(-1*alPlace) ) )
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_MathLog()
'Overview                    : 引数を底とする対数
'Detailed Description        : 工事中
'Argument
'     adbBase                : 底
'     adbAntilogarithm       : 真数
'Return Value
'     冪指数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_MathLog( _
    byVal adbBase _
    , byVal adbAntilogarithm _
    )
    func_MathLog = log(adbAntilogarithm)/log(adbBase)
End Function


'###################################################################################################
'ユーティリティ系の関数
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : util_escapeForPs()
'Overview                    : powershell用の特殊文字エスケープを行う
'Detailed Description        : 公式サイトに書いていないがshellから起動する場合にいくつか対応しないと動作しないものがある
'                              https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_special_characters?view=powershell-7.4
'Argument
'     asTgt                  : 対象
'Return Value
'     対象にエスケープ処理した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_escapeForPs( _
    byVal asTgt _
    )
    Const Cs_BACKQUOTE = "`"
    Dim vLst : vLst = Array("(",")"," ")
    
    Dim i, sRet : sRet = asTgt
    For Each i In vLst
        sRet = Replace(sRet, i, Cs_BACKQUOTE&i)
    Next
    util_escapeForPs = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : util_getIpAddress()
'Overview                    : 自身のIPアドレスを取得する
'Detailed Description        : IPアドレスを格納したオブジェクトを返す
'Argument
'     なし
'Return Value
'     IPアドレスを格納したオブジェクトの配列
'                              内容は以下のとおり
'                               Key             Value                   例
'                               --------------  ----------------------  ----------------------------
'                               "Caption"       Adapter名               -
'                               "Ip"            以下オブジェクト        -
'                              
'                              IP Addressを格納したオブジェクト
'                               Key             Value                   例
'                               --------------  ----------------------  ----------------------------
'                               "V4"            IP Address(v4)          192.168.11.52
'                               "V6"            IP Address(v6)          fe80::ba87:1e93:59ab:28f7%18
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/10         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_getIpAddress( _
    )
    Dim sMyComp, oAdapter, oAddress, oRet, oIpv4, oIpv6
    
    For Each oAdapter in CreateObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
         For Each oAddress in oAdapter.IPAddress
             If new_ArrSplit(oAddress, ".").length=4 Then
             'IPv4
                 cf_bind oIpv4, oAddress
             Else
             'IPv6
                 cf_bind oIpv6, oAddress
             End If
         Next
         cf_push oRet, new_DicOf(Array("Caption", oAdapter.Caption, "Ip", new_DicOf(Array("V4", oIpv4, "V6", oIpv6))))
    Next
    util_getIpAddress = oRet
    
    Set oAddress = Nothing
    Set oAdapter = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : util_isZipWithPassword()
'Overview                    : パスワード付きzipファイルか判定する
'Detailed Description        : https://vbavb.com/zip/
'Argument
'     asPath                 : パス
'Return Value
'     結果 True:パスワード付きzipファイル / False:パスワード付きzipファイルでない
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/01/29         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_isZipWithPassword( _
    byVal asPath _
    )
    util_isZipWithPassword = False
    If Not new_Fso().FileExists(asPath) Then Exit Function
    With new_Adodb()
        .Type = 1       'adTypeBinary
        .Open
        .LoadFromFile asPath
        .Position = 6
        If Hex(AscB(.Read(1)))=1 Then util_isZipWithPassword = True
        .Close
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : util_randStr()
'Overview                    : ランダムな文字列を生成する
'Detailed Description        : 指定した文字（配列）、指定した回数でランダムな文字列を生成する
'Argument
'     avStrings              : 文字の配列
'     alTimes                : 回数
'Return Value
'     生成した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_randStr( _
    byRef avStrings _
    , byVal alTimes _
    )
    Dim lPos, sRet, lUb
    sRet = "" : lUb = Ubound(avStrings)
    For lPos = 1 To alTimes
        sRet = sRet & avStrings( math_rand(0, lUb, 0) )
    Next
    util_randStr = sRet
End Function

'***************************************************************************************************
'Function/Sub Name           : util_unzip()
'Overview                    : zipファイルを解凍する
'Detailed Description        : https://excel-vba.work/2021/12/10/%E3%80%90vba%E3%80%91zip%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%E8%A7%A3%E5%87%8D%E5%B1%95%E9%96%8B%E3%81%99%E3%82%8B/#google_vignette
'Argument
'     asPath                 : zipファイルのパス
'     asDestination          : 解凍先
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/03         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_unzip( _
    byVal asPath _
    , byVal asDestination _
    )
    util_unzip=False
    'PowerShellのコマンドを作成
    Dim sCmd : sCmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Force" _
        & " -Path " & fs_wrapInQuotes(util_escapeForPs(asPath)) _
        & " -DestinationPath " & fs_wrapInQuotes(util_escapeForPs(asDestination))
    '作成したコマンドをサイレント実行する
    util_unzip = fw_runShellSilently(sCmd)
End Function

'***************************************************************************************************
'Function/Sub Name           : util_zip()
'Overview                    : zipファイルを作成する
'Detailed Description        : https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.archive/compress-archive?view=powershell-7.4
'Argument
'     asPath                 : 圧縮するファイルのパス
'     asDestination          : zipファイルのパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/05         Y.Fujii                  First edition
'***************************************************************************************************
Private Function util_zip( _
    byRef asPath _
    , byVal asDestination _
    )
    util_zip=False
    If new_Fso().FileExists(asDestination) Then Exit Function

    '圧縮するファイルのパスを連結
    Dim sPath : sPath = new_ArrOf(asPath).map(new_Func("(e,i,a)=>fs_wrapInQuotes(util_escapeForPs(e))")).join(",")

    'PowerShellのコマンドを作成
    Dim sCmd : sCmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive" _
        & " -Path " & sPath _
        & " -DestinationPath " & fs_wrapInQuotes(util_escapeForPs(asDestination))

    '作成したコマンドをサイレント実行する
    util_zip = fw_runShellSilently(sCmd)
End Function

'###################################################################################################
'ファイル操作系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : fs_copyFile()
'Overview                    : ファイルをコピーする
'Detailed Description        : FileSystemObjectのCopyFile()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asFrom                 : コピー元ファイルのフルパス
'     asTo                   : コピー先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/24         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyFile( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_copyFile = func_FsGeneralExecutor(Array(asFrom, asTo), "CopyFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_copyFolder()
'Overview                    : フォルダをコピーする
'Detailed Description        : FileSystemObjectのCopyFolder()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asFrom                 : コピー元フォルダのフルパス
'     asTo                   : コピー先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyFolder( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_copyFolder = func_FsGeneralExecutor(Array(asFrom, asTo), "CopyFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_copyHere()
'Overview                    : フォルダまたはファイルをフォルダーにコピーする
'Detailed Description        : Windowsシェル関数のfolder.Copyhere()準拠
'                              https://learn.microsoft.com/ja-jp/windows/win32/shell/folder-copyhere
'                              コピー先フォルダが存在しない場合は作成する
'Argument
'     asPath                 : コピーするフォルダまたはファイルのパス
'     asDestination          : コピー先フォルダのパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/18         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_copyHere( _
    byVal asPath _
    , byVal asDestination _
    )

    On Error Resume Next
    'コピーするフォルダまたはファイルの存在確認
    new_FolderItem2Of(asPath)
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If

    'コピー
    new_ShellApp().Namespace(asDestination).CopyHere asPath
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If
    On Error Goto 0

    'コピー先の存在確認
    With new_Fso()
        Dim sPath : sPath = .BuildPath(asDestination, .GetFileName(asPath))
    End With
    On Error Resume Next
    new_FolderItem2Of(sPath)
    If Err.Number<>0 Then
        Set fs_copyHere = new_Ret(False)
        Err.Clear
        Exit Function
    End If
    On Error Goto 0

    Set fs_copyHere = new_Ret(True)

End Function

'***************************************************************************************************
'Function/Sub Name           : fs_createFolder()
'Overview                    : フォルダを作成する
'Detailed Description        : FileSystemObjectのCreateFolder()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asPath                 : 作成するフォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_createFolder( _
    byVal asPath _
    )
    Set fs_createFolder = func_FsGeneralExecutor(Array(asPath), "CreateFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_deleteFile()
'Overview                    : ファイルを削除する
'Detailed Description        : FileSystemObjectのDeleteFile()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asPath                 : 削除するファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_deleteFile( _
    byVal asPath _
    )
    Set fs_deleteFile = func_FsGeneralExecutor(Array(asPath), "DeleteFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_deleteFolder()
'Overview                    : フォルダを削除する
'Detailed Description        : FileSystemObjectのDeleteFolder()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asPath                 : 削除するフォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_deleteFolder( _
    byVal asPath _
    )
    Set fs_deleteFolder = func_FsGeneralExecutor(Array(asPath), "DeleteFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_moveFile()
'Overview                    : ファイルを移動する
'Detailed Description        : FileSystemObjectのMoveFile()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asFrom                 : 移動元ファイルのフルパス
'     asTo                   : 移動先ファイルのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_moveFile( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_moveFile = func_FsGeneralExecutor(Array(asFrom, asTo), "MoveFile")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_moveFolder()
'Overview                    : フォルダを移動する
'Detailed Description        : FileSystemObjectのMoveFolder()と同等
'                              func_FsGeneralExecutor()に委譲する
'Argument
'     asFrom                 : 移動元フォルダのフルパス
'     asTo                   : 移動先フォルダのフルパス
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_moveFolder( _
    byVal asFrom _
    , byVal asTo _
    )
    Set fs_moveFolder = func_FsGeneralExecutor(Array(asFrom, asTo), "MoveFolder")
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_readFile()
'Overview                    : Unicode形式のファイルを読んで中身を取得する
'Detailed Description        : func_FsReadFile()に委譲し以下の設定で読込む
'                               ファイルの形式         ：Unicode形式
'Argument
'     asPath                 : 入力先のフルパス
'Return Value
'     ファイルの内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_readFile( _
    byVal asPath _
    )
    Set fs_readFile = func_FsReadFile(asPath, -1)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_wrapInQuotes()
'Overview                    : 引用符（"：ダブルクォーテーション）で囲む
'Detailed Description        : 対象に引用符（"：ダブルクォーテーション）を含む場合はエスケープする
'Argument
'     asTgt                  : 対象
'Return Value
'     対象を引用符（"：ダブルクォーテーション）で囲んだ文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2024/02/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_wrapInQuotes( _
    byVal asTgt _
    )
    fs_wrapInQuotes = Chr(34) & Replace(asTgt, Chr(34), Chr(34)&Chr(34)) & Chr(34)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_writeFile()
'Overview                    : Unicode形式でファイル出力する
'Detailed Description        : func_FsWriteFile()に委譲し以下の設定で出力する
'                               出力モード            ：既存のファイルを新しいデータで置き換える
'                               ファイルが存在しない場合：新しいファイルを作成する
'                               ファイルの形式         ：Unicode形式
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_writeFile( _
    byVal asPath _
    , byVal asCont _
    )
    Set fs_writeFile = func_FsWriteFile(asPath, 2, True, -1, asCont)
End Function

'***************************************************************************************************
'Function/Sub Name           : fs_writeFileDefault()
'Overview                    : システムの既定の形式でファイル出力する
'Detailed Description        : func_FsWriteFile()に委譲し以下の設定で出力する
'                               出力モード            ：既存のファイルを新しいデータで置き換える
'                               ファイルが存在しない場合：新しいファイルを作成する
'                               ファイルの形式         ：システムの既定
'Argument
'     asPath                 : 出力先のフルパス
'     asCont                 : 出力する内容
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_writeFileDefault( _
    byVal asPath _
    , byVal asCont _
    )
    Set fs_writeFileDefault = func_FsWriteFile(asPath, 2, True, -2, asCont)
End Function


'***************************************************************************************************
'Function/Sub Name           : fs_getAllFiles()
'Overview                    : フォルダ配下のファイルオブジェクトを取得する
'Detailed Description        : 工事中
'Argument
'     asPath                 : ファイル/フォルダのパス
'Return Value
'     Fileオブジェクト相当（アダプターでラップした）のオブジェクトの配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function fs_getAllFiles( _
    byVal asPath _
    )
    fs_getAllFiles = func_FsGetAllFilesByFso(asPath)
'    fs_getAllFiles = func_FsGetAllFilesByShell(asPath)
'    fs_getAllFiles = func_FsGetAllFilesByDir(asPath)
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByFso()
'Overview                    : フォルダ配下のファイルオブジェクトを取得する（FSO版）
'Detailed Description        : zipファイル内の検索はfunc_FsGetAllFilesByShell()に委譲する
'Argument
'     asPath                 : ファイル/フォルダのパス
'Return Value
'     Fileオブジェクト相当（アダプターでラップした）のオブジェクトの配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByFso( _
    byVal asPath _
    )
    If new_Fso().FolderExists(asPath) Then
    'フォルダの場合
        Dim oFolder : Set oFolder = new_FolderOf(asPath)
        Dim oEle, vRet()
        'ファイルの取得
        For Each oEle In oFolder.Files
            If StrComp(new_Fso().GetExtensionName(oEle.Path), "zip", vbTextCompare)=0 Then
            'zipファイルの場合、func_FsGetAllFilesByShell()でzip内のファイルリストを取得する
                cf_pushA vRet, func_FsGetAllFilesByShell(oEle.Path)
            Else
            'zipファイル以外の場合、ファイル情報を取得する
                cf_push vRet, new_AdptFileOf(oEle.Path)
            End If
        Next
        'フォルダの取得
        For Each oEle In oFolder.SubFolders
            cf_pushA vRet, func_FsGetAllFilesByFso(oEle)
        Next
        func_FsGetAllFilesByFso = vRet
    Else
    'ファイルの場合
        func_FsGetAllFilesByFso = Array(new_AdptFileOf(asPath))
    End If

    Set oFolder = Nothing
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByShell()
'Overview                    : フォルダ配下のファイルオブジェクトを取得する（ShellApp版）
'Detailed Description        : zipファイル内のファイルリストを取得できる
'Argument
'     asPath                 : ファイル/フォルダのパス
'Return Value
'     Fileオブジェクト相当（アダプターでラップした）のオブジェクトの配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByShell( _
    byVal asPath _
    )
    '処理タイプ判定
    Dim boFlg : boFlg = True 'AsFolder
    If new_Fso().FileExists(asPath) Then
        If StrComp(new_Fso().GetExtensionName(asPath), "zip", vbTextCompare)<>0 Then boFlg=False 'AsFile
    End If
    
    If boFlg Then
    'フォルダかzipファイルの場合
        Dim oFolder : Set oFolder = new_ShellApp().Namespace(asPath)
        Dim oItem, vRet()
        For Each oItem In oFolder.Items
        'フォルダ内全てのアイテムについて
            If oItem.IsFolder Then
            'フォルダの場合
                cf_pushA vRet, func_FsGetAllFilesByShell(oItem.Path)
            Else
            'ファイルの場合
                cf_push vRet, new_AdptFile(oItem)
            End If
        Next
        func_FsGetAllFilesByShell = vRet
        Set oItem = Nothing
    Else
    '上記以外の場合
        func_FsGetAllFilesByShell = Array(new_AdptFileOf(asPath))
    End If
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGetAllFilesByDir()
'Overview                    : フォルダ配下のファイルオブジェクトを取得する（Dir版）
'Detailed Description        : zipファイル内の検索はfunc_FsGetAllFilesByShell()に委譲する
'Argument
'     asPath                 : ファイル/フォルダのパス
'Return Value
'     Fileオブジェクト相当（アダプターでラップした）のオブジェクトの配列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/11/25         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGetAllFilesByDir( _
    byVal asPath _
    )
    Dim sDir : sDir = "dir /S /B /A-D " & fs_wrapInQuotes(asPath)
    Dim sTmpPath : sTmpPath = fw_getTempPath()
    fw_runShellSilently "cmd /U /C " & sDir & " > " & fs_wrapInQuotes(sTmpPath)
    Dim sLists : sLists = fs_readFile(sTmpPath)
    fs_deleteFile sTmpPath
    
    Dim vArrList : vArrList = Split(sLists, vbNewLine)
    Redim Preserve vArrList(Ubound(vArrList)-1)
    Dim sList, vRet()
    For Each sList In vArrList
        If StrComp(new_Fso().GetExtensionName(sList), "zip", vbTextCompare)=0 Then
        'zipファイルの場合、func_FsGetAllFilesByShell()でzip内のファイルリストを取得する
            cf_pushA vRet, func_FsGetAllFilesByShell(sList)
        Else
        'zipファイル以外の場合、ファイル情報を取得する
            cf_push vRet, new_AdptFileOf(sList)
        End If
    Next
    func_FsGetAllFilesByDir = vRet
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsGeneralExecutor()
'Overview                    : Fsoコマンド汎用実行関数
'Detailed Description        : 工事中
'Argument
'     asPath                 : パス
'     asCmd                  : 実行コマンド
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/30         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsGeneralExecutor( _
    byRef asPath _
    , byVal asCmd _
    )
    With new_Fso()
        On Error Resume Next
        Select Case asCmd
            Case "CopyFile":
                .CopyFile asPath(0), asPath(1)
            Case "CopyFolder":
                .CopyFolder asPath(0), asPath(1)
            Case "CreateFolder":
                .CreateFolder asPath(0)
            Case "DeleteFile":
                .DeleteFile asPath(0)
            Case "DeleteFolder":
                .DeleteFolder asPath(0)
            Case "MoveFile":
                .MoveFile asPath(0), asPath(1)
            Case "MoveFolder":
                .MoveFolder asPath(0), asPath(1)
            Case Else
                Err.Raise 9999, "libCom.vbs:func_FsGeneralExecutor()", "不正な実行コマンド："&asCmd
        End Select
        Set func_FsGeneralExecutor = new_RetByState(True,False)
        On Error Goto 0
    End With
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsReadFile()
'Overview                    : ファイルを読んで中身を取得する
'Detailed Description        : 工事中
'Argument
'     asPath                 : 入力先のフルパス
'     alFormat               : ファイルの形式
'                               -2：システムの既定 / -1：Unicode / 0：Ascii
'Return Value
'     ファイルの内容
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsReadFile( _
    byVal asPath _
    , byVal alFormat _
    )
    Dim sRet : sRet = Empty
    On Error Resume Next
    With new_Ts(asPath, 1, False, alFormat)
        sRet = .ReadAll
        .Close
    End With
    Set func_FsReadFile = new_Ret(sRet)
    On Error Goto 0
End Function

'***************************************************************************************************
'Function/Sub Name           : func_FsWriteFile()
'Overview                    : ファイルに出力する
'Detailed Description        : 工事中
'Argument
'     asPath                 : 出力先のフルパス
'     alMode                 : 出力モード
'                               2：既存のファイルを新しいデータで置き換える / 8：ファイルの最後に書き込み）
'     aboCreate              : ファイルが存在しない場合に新しいファイルを作成できるかどうかを示す
'                               True：新しいファイルを作成する / False：作成しない
'     alFormat               : ファイルの形式
'                               -2：システムの既定 / -1：Unicode / 0：Ascii
'     asCont                 : 出力する内容
'Return Value
'     結果 True:成功 / False:失敗
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/12/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_FsWriteFile( _
    byVal asPath _
    , byVal alMode _
    , byVal aboCreate _
    , byVal alFormat _
    , byVal asCont _
    )
    On Error Resume Next
    With new_Ts(asPath, alMode, aboCreate, alFormat)
        .Write asCont
        .Close
    End With
    Set func_FsWriteFile=new_RetByState(True,False)
    On Error Goto 0
End Function





'###################################################################################################
'エクセル系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : sub_CM_ExcelSaveAs()
'Overview                    : エクセルファイルを別名で保存して閉じる
'Detailed Description        : 工事中
'Argument
'     aoWorkBook             : エクセルのワークブック
'     asPath                 : 保存するファイルのフルパス
'     alFileformat           : XlFileFormat 列挙体（デフォルトはxlOpenXMLWorkbook 51 Excelブック）
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_CM_ExcelSaveAs( _
    byRef aoWorkBook _
    , byVal asPath _
    , byVal alFileformat _
    )
    If Not(IsNumeric(alFileformat)) Then
        alFileformat = 51                  'xlOpenXMLWorkbook 51 Excelブック
    End If
    Call aoWorkBook.SaveAs( _
                            asPath _
                            , alFileformat _
                            , , _
                            , False _
                            , False _
                            )
    Call aoWorkBook.Close(False)
End Sub

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelOpenFile()
'Overview                    : エクセルファイルを読み取り専用／ダイアログなしで開く
'Detailed Description        : 工事中
'Argument
'     aoExcel                : エクセル
'     asPath                 : エクセルファイルのフルパス
'Return Value
'     開いたエクセルのワークブック
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelOpenFile( _
    byRef aoExcel _
    , byVal asPath _
    )    
    Set func_CM_ExcelOpenFile = aoExcel.Workbooks.Open( _
                                                        asPath _
                                                        , 0 _
                                                        , True _
                                                        , , , _
                                                        , True _
                                                        )
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_ExcelGetTextFromAutoshape()
'Overview                    : エクセルのオートシェイプのテキストを取り出す
'Detailed Description        : エラーは無視する
'Argument
'     aoAutoshape            : オートシェイプ
'Return Value
'     オートシェイプのテキスト
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/09/27         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ExcelGetTextFromAutoshape( _
    byRef aoAutoshape _
    )
    func_CM_ExcelGetTextFromAutoshape = aoAutoshape.TextFrame.Characters.Text
End Function


'###################################################################################################
'文字列操作系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConvOnlyAlphabet()
'Overview                    : 英字だけ大文字／小文字に変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類 1:UpperCase,2:LowerCase
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConvOnlyAlphabet( _
    byVal asTarget _
    , byVal alConversion _
    )
    Dim sChar, asTargetTmp
    
    '1文字ずつ判定する
    Dim asTargetNew : asTargetNew = asTarget
    Dim lPos : lPos = 1
    Do While Len(asTargetNew) >= lPos
        '変換対象の1文字を取得
        sChar = Mid(asTargetNew, lPos, 1)
        
        If new_Char().whatType(sChar)<3 Then
        '変換対象が英字の場合のみ変換する
            asTargetTmp = ""
            
            '変換対象の文字までの文字列
            If lPos > 1 Then
                asTargetTmp = Mid(asTargetNew, 1, lPos-1)
            End If
            
            '変換した1文字を結合
            sChar = func_CM_StrConv(sChar, alConversion)
            asTargetTmp = asTargetTmp & sChar
            
            '変換対象の文字移行の文字列を結合
            If lPos < Len(asTargetNew) Then
                asTargetTmp = asTargetTmp & Mid(asTargetNew, lPos+1, Len(asTargetNew)-lPos)
            End If
            
            '変換後の文字列を格納
            asTargetNew = asTargetTmp
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrConvOnlyAlphabet = asTargetNew
    
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrConv()
'Overview                    : 文字列を指定のとおり変換する
'Detailed Description        : 工事中
'Argument
'     asTarget               : 変換する文字列
'     alConversion           : 実行する変換の種類（現時点で1,2のみ実装）
'                                 1:文字列を大文字に変換
'                                 2:文字列を小文字に変換
'                                 3:文字列内のすべての単語の最初の文字を大文字に変換
'                                 4:文字列内の狭い (1 バイト) 文字をワイド (2 バイト) 文字に変換
'                                 8:文字列内のワイド (2 バイト) 文字を狭い (1 バイト) 文字に変換
'                                16:文字列内のひらがな文字をカタカナ文字に変換
'                                32:文字列内のカタカナ文字をひらがな文字に変換
'Return Value
'     変換した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrConv( _
    byVal asTarget _
    , byVal alalConversion _
    )
    Dim sChar, asTargetTmp
    func_CM_StrConv = asTarget
    Select Case alalConversion
        Case 1:
            func_CM_StrConv = UCase(asTarget)
        Case 2:
            func_CM_StrConv = LCase(asTarget)
    End Select
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_StrLen()
'Overview                    : 全角は2文字、半角は1文字として文字数を返す
'Detailed Description        : 工事中
'Argument
'     asTarget               : 文字列
'Return Value
'     文字数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_StrLen( _
    byVal asTarget _
    )
    '1文字ずつ判定する
    Dim sChar
    Dim lLength : lLength = 0
    Dim lPos : lPos = 1
    Do While Len(asTarget) >= lPos
        '1文字を取得
        sChar = Mid(asTarget, lPos, 1)
        
        If (Asc(sChar) And &HFF00) <> 0 Then
            lLength = lLength+2
        Else
            lLength = lLength+1
        End If
        
        'カウントアップ
        lPos = lPos+1
    Loop
    
    func_CM_StrLen = lLength
End Function


'###################################################################################################
'配列系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_ArrayGetDimensionNumber()
'Overview                    : 配列の次元数を求める
'Detailed Description        : 工事中
'Argument
'     avArray                : 配列
'Return Value
'     配列の次元数
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/11/19         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_ArrayGetDimensionNumber( _
    byRef avArray _ 
    )
   If Not IsArray(avArray) Then Exit Function
   On Error Resume Next
   Dim lNum : lNum = 0
   Dim lTemp
   Do
       lNum = lNum + 1
       lTemp = UBound(avArray, lNum)
   Loop Until Err.Number <> 0
   Err.Clear
   func_CM_ArrayGetDimensionNumber = lNum - 1
End Function


'###################################################################################################
'その他
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_FillInTheCharacters()
'Overview                    : 文字を埋める
'Detailed Description        : 対象文字の不足桁を指定したアライメントで指定した文字の1文字目で埋める
'                              対象文字に不足桁がない場合は、指定した文字数で切り取る
'Argument
'     asTarget               : 対象文字列
'     alWordCount            : 文字数
'     asToFillCharacter      : 埋める文字
'     aboIsCutOut            : 文字数で切り取り（True：する/False：しない）
'     aboIsRightAlignment    : アライメント（True：右寄せ/False：左寄せ）
'Return Value
'     埋めた文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/01/04         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FillInTheCharacters( _
    byVal asTarget _
    , byVal alWordCount _
    , byVal asToFillCharacter _
    , byVal aboIsCutOut _
    , byVal aboIsRightAlignment _
    )
    
    '切り取りなしで対象文字列が文字数より大きい場合は処理を抜ける
    Dim lTargetLen : lTargetLen = Len(asTarget)
    If Not(aboIsCutOut) And lTargetLen>=alWordCount Then
        func_CM_FillInTheCharacters = asTarget
        Exit Function
    End If
    
    '埋める文字列の作成
    Dim sFillStrings : sFillStrings = ""
    If alWordCount-lTargetLen > 0 Then
        sFillStrings = String(alWordCount-lTargetLen , asToFillCharacter)
    End If
    
    Dim sResult
    'アライメント指定によって文字列を埋める
    If aboIsRightAlignment Then
        sResult = sFillStrings & asTarget
    Else
        sResult = asTarget & sFillStrings
    End If
    
    '切り取りなしの場合は処理を抜ける
    If Not(aboIsCutOut) Then
        func_CM_FillInTheCharacters = sResult
        Exit Function
    End If
    
    'アライメント指定によって文字列を切り取る
    If aboIsRightAlignment Then
        sResult = Right(sResult, alWordCount)
    Else
        sResult = Left(sResult, alWordCount)
    End If
    func_CM_FillInTheCharacters = sResult
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_FormatDecimalNumber()
'Overview                    : 浮動小数点型を整形する
'Detailed Description        : 工事中
'Argument
'     adbNumber              : 浮動小数点型の数値
'     alDecimalPlaces        : 小数の桁数
'Return Value
'     整形した浮動小数点型
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/08/26         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FormatDecimalNumber( _
    byVal adbNumber _
    , byVal alDecimalPlaces _
    )
    func_CM_FormatDecimalNumber = Fix(adbNumber) & "." _
                             & func_CM_FillInTheCharacters( _
                                                          Abs(Fix( (adbNumber - Fix(adbNumber))*10^alDecimalPlaces )) _
                                                          , alDecimalPlaces _
                                                          , "0" _
                                                          , False _
                                                          , True _
                                                          )
End Function


'###################################################################################################
'ユーティリティ系
'###################################################################################################

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilSortDefaultFunc()
'Overview                    : 要素の比較結果を返す
'Detailed Description        : ソート関数群で使うデフォルトの関数
'Argument
'     aoCurrentValue         : 配列の要素
'     aoNextValue            : 次の配列の要素
'Return Value
'     要素の比較結果
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/15         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilSortDefaultFunc( _
    byRef aoCurrentValue _
    , byRef aoNextValue _
    )
    func_CM_UtilSortDefaultFunc = aoCurrentValue>aoNextValue
End Function

'***************************************************************************************************
'Function/Sub Name           : func_CM_UtilJoin()
'Overview                    : Join関数
'Detailed Description        : vbscriptのJoin関数と同等の機能
'Argument
'     avArr                  : 配列
'     asDel                  : 区切り文字
'Return Value
'     配列の各要素を連結した文字列
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/17         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_UtilJoin( _
    byRef avArr _
    , byVal asDel _
    )
    func_CM_UtilJoin = Join(avArr, asDel)
End Function
