'***************************************************************************************************
'FILENAME                    : clsCmCharacterType.vbs
'Overview                    : 文字種類管理クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/28         Y.Fujii                  First edition
'***************************************************************************************************
Class clsCmCharacterType
    'クラス内変数、定数
    Private Cl_MAX_POWER_OF_2
    Private PvSettings
    Private PoChar2Type
    Private PoType2Chars
    
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
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Initialize()
        Cl_MAX_POWER_OF_2 = 16          '2^16 = 65536 <= Typeの最大値
        PvSettings = Array( _
              Array( Array("A", "Z") ) _
              , Array( Array("a", "z") ) _
              , Array( Array("0", "9") ) _
              , Array( Array(" ", "/"), Array(":", "@"), Array("[", "`"), Array("{", "~") ) _
              , Array( Array("ｦ", "ｯ"), Array("ｱ", "ﾟ") ) _
              , Array( Array("｡", "･"), Array("ｰ", "ｰ") ) _
              , Array( Array("Ａ", "Ｚ") ) _
              , Array( Array("ａ", "ｚ") ) _
              , Array( Array("０", "９") ) _
              , Array( Array("　", "〓"), Array("∈", "∩"), Array("∧", "∃"), Array("∠", "∬"), Array("Å", "¶"), Array("◯", "◯") ) _
              , Array( Array("ぁ", "ん") ) _
              , Array( Array("ァ", "ヶ") ) _
              , Array( Array("Α", "Ω"), Array("А", "Я") ) _
              , Array( Array("α", "ω"), Array("а", "я") ) _
              , Array( Array("─", "╂") ) _
              , Array( Array("亜", "腕") ) _
              , Array( Array("弌", "滌"), Array("漾", "熙") ) _
              )
        Set PoChar2Type = new_Dic()
        Set PoType2Chars = new_Dic()
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
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoChar2Type = Nothing
        Set PoType2Chars = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : whatType()
    'Overview                    : 文字の種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     asChar                 : 文字
    'Return Value
    '     文字の種類（内容はgetCharList()の引数（alType）と同じ）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function whatType( _
        byVal asChar _
        )
        Dim bCode : bCode = Asc(asChar)
        If PoChar2Type.Exists(bCode) Then
            whatType = PoChar2Type.Item(bCode)
            Exit Function
        End If

        Dim lPowerOf2 : lPowerOf2 = 0
        Do While lPowerOf2 <= Cl_MAX_POWER_OF_2
            If Not PoType2Chars.Exists(2^lPowerOf2) Then
                sub_CmCharTypeCreateDefinitionsByCharacterType lPowerOf2
                If PoChar2Type.Exists(bCode) Then
                    whatType = PoChar2Type.Item(bCode)
                    Exit Function
                End If
            End If
            lPowerOf2 = lPowerOf2+1
        Loop
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : getCharList()
    'Overview                    : 指定した文字の種類の配列を返す
    'Detailed Description        : http://charset.7jp.net/sjis.html
    'Argument
    '     alType                 : 文字の種類（複数指定する場合は以下の和を設定する）
    '                                    1:半角英字大文字
    '                                    2:半角英字小文字
    '                                    4:半角数字
    '                                    8:半角記号
    '                                   16:半角カタカナ
    '                                   32:半角カタカナ記号
    '                                   64:全角英字大文字
    '                                  128:全角英字小文字
    '                                  256:全角数字
    '                                  512:全角記号
    '                                 1024:全角ひらがな
    '                                 2048:全角カタカナ
    '                                 4096:全角ギリシャ、キリル文字の大文字
    '                                 8192:全角ギリシャ、キリル文字の小文字
    '                                16384:全角線枠
    '                                32768:全角漢字 第1水準(16区～47区)
    '                                65536:全角漢字 第2水準(48区～84区)
    'Return Value
    '     文字の種類（配列）
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function getCharList( _
        byVal alType _
        )
        Dim lType : lType = alType
        Dim lPowerOf2 : lPowerOf2 = Cl_MAX_POWER_OF_2
        Dim vRet : Set vRet = new_Arr()
        Dim lQuotient,lDivide
        Do Until lPowerOf2<0
            lDivide = 2^lPowerOf2
            lQuotient = lType \ lDivide
            lType = lType Mod lDivide
            If lQuotient>0 Then
                If Not PoType2Chars.Exists(lDivide) Then
                    sub_CmCharTypeCreateDefinitionsByCharacterType lPowerOf2
                End If
                vRet.pushMulti PoType2Chars.Item(lDivide)
            End If
            lPowerOf2 = lPowerOf2 - 1
        Loop
        getCharList = vRet.items()
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : sub_CmCharTypeCreateDefinitionsByCharacterType
    'Overview                    : 指定した文字種類の定義を作成する
    'Detailed Description        : 工事中
    'Argument
    '     alPowerOf2             : 文字の種類（内容はgetCharList()の引数（alType）と同じ）を2^nとした場合のn
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'Histroy
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub sub_CmCharTypeCreateDefinitionsByCharacterType( _
        byVal alPowerOf2 _
        )
        Dim lType : lType = 2^alPowerOf2
        If PoType2Chars.Exists(lType) Then Exit Sub

        Dim vArr : vArr = Array()
        Dim vSetting : vSetting = PvSettings(alPowerOf2)
        Dim vEle, bCode, sCodeHex
        For Each vEle In vSetting
            For bCode = Asc(vEle(0)) To Asc(vEle(1))
                sCodeHex = "" : If bCode<0 Then sCodeHex = Right(Hex(bCode),2)
                If bCode>=0 Or (sCodeHex<>"7F" And "3F"<sCodeHex And sCodeHex<"FD" ) Then
                    PoChar2Type.Add bCode, lType
                    cf_push vArr, Chr(bCode)
                End If
            Next
        Next
        PoType2Chars.Add lType, vArr
    End Sub
    
End Class
