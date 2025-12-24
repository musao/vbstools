'***************************************************************************************************
'FILENAME                    : CharacterType.vbs
'Overview                    : 文字種類管理クラス
'Detailed Description        : 工事中
'---------------------------------------------------------------------------------------------------
'History
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2023/10/28         Y.Fujii                  First edition
'***************************************************************************************************
Class CharacterType
    'クラス内変数、定数
    Private Cl_MAX_POWER_OF_2, PvSettings, PoChar2Type, PoType2Chars
    
    '***************************************************************************************************
    'Function/Sub Name           : Class_Initialize()
    'Overview                    : コンストラクタ
    'Detailed Description        : 内部変数の初期化
    'Argument
    '     なし
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub Class_Terminate()
        Set PoChar2Type = Nothing
        Set PoType2Chars = Nothing
    End Sub
    
    '***************************************************************************************************
    'Function/Sub Name           : Property Get type*()
    'Overview                    : 文字の種類の値を返す
    'Detailed Description        : 工事中
    'Argument
    '     なし
    'Return Value
    '     文字の種類の値
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/11/03         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Property Get typeHalfWidthAlphabetUppercase()
        typeHalfWidthAlphabetUppercase = 2^0
    End Property
    Public Property Get typeHalfWidthAlphabetLowercase()
        typeHalfWidthAlphabetLowercase = 2^1
    End Property
    Public Property Get typeHalfWidthNumbers()
        typeHalfWidthNumbers = 2^2
    End Property
    Public Property Get typeHalfWidthSymbol()
        typeHalfWidthSymbol = 2^3
    End Property
    Public Property Get typeHalfWidthKatakana()
        typeHalfWidthKatakana = 2^4
    End Property
    Public Property Get typeHalfWidthKatakanaSymbol()
        typeHalfWidthKatakanaSymbol = 2^5
    End Property
    Public Property Get typeFullWidthAlphabetUppercase()
        typeFullWidthAlphabetUppercase = 2^6
    End Property
    Public Property Get typeFullWidthAlphabetLowercase()
        typeFullWidthAlphabetLowercase = 2^7
    End Property
    Public Property Get typeFullWidthNumbers()
        typeFullWidthNumbers = 2^8
    End Property
    Public Property Get typeFullWidthSymbol()
        typeFullWidthSymbol = 2^9
    End Property
    Public Property Get typeFullWidthHiragana()
        typeFullWidthHiragana = 2^10
    End Property
    Public Property Get typeFullWidthKatakana()
        typeFullWidthKatakana = 2^11
    End Property
    Public Property Get typeFullWidthGreekCyrillicUppercase()
        typeFullWidthGreekCyrillicUppercase = 2^12
    End Property
    Public Property Get typeFullWidthGreekCyrillicLowercase()
        typeFullWidthGreekCyrillicLowercase = 2^13
    End Property
    Public Property Get typeFullWidthLineFrame()
        typeFullWidthLineFrame = 2^14
    End Property
    Public Property Get typeFullWidthKanjiLevel1()
        typeFullWidthKanjiLevel1 = 2^15
    End Property
    Public Property Get typeFullWidthKanjiLevel2()
        typeFullWidthKanjiLevel2 = 2^16
    End Property
    '全て
    Public Property Get typeAll()
        Dim i,lStt,lEnd,lRet : lRet = 0
        lStt = 0
        lEnd = Cl_MAX_POWER_OF_2
        For i=lStt To lEnd
            lRet = lRet + 2^i
        Next
        typeAll = lRet
    End Property
    '半角のグループ
    Public Property Get typeHalfWidthAlphanumeric()
        Dim i,lStt,lEnd,lRet : lRet = 0
        lStt = math_log2(typeHalfWidthAlphabetUppercase)
        lEnd = math_log2(typeHalfWidthNumbers)
        For i=lStt To lEnd
            lRet = lRet + 2^i
        Next
        typeHalfWidthAlphanumeric = lRet
    End Property
    Public Property Get typeHalfWidthAlphanumericSymbols()
        Dim i,lStt,lEnd,lRet : lRet = 0
        lStt = math_log2(typeHalfWidthAlphabetUppercase)
        lEnd = math_log2(typeHalfWidthSymbol)
        For i=lStt To lEnd
            lRet = lRet + 2^i
        Next
        typeHalfWidthAlphanumericSymbols = lRet
    End Property
    Public Property Get typeHalfWidthAll()
        Dim i,lStt,lEnd,lRet : lRet = 0
        lStt = math_log2(typeHalfWidthAlphabetUppercase)
        lEnd = math_log2(typeHalfWidthKatakanaSymbol)
        For i=lStt To lEnd
            lRet = lRet + 2^i
        Next
        typeHalfWidthAll = lRet
    End Property
    '全角のグループ
    Public Property Get typeFullWidthAll()
        Dim i,lStt,lEnd,lRet : lRet = 0
        lStt = math_log2(typeFullWidthAlphabetUppercase)
        lEnd = math_log2(typeFullWidthKanjiLevel2)
        For i=lStt To lEnd
            lRet = lRet + 2^i
        Next
        typeFullWidthAll = lRet
    End Property
    
    '***************************************************************************************************
    'Function/Sub Name           : whatType()
    'Overview                    : 文字の種類を返す
    'Detailed Description        : 工事中
    'Argument
    '     asChar                 : 文字
    'Return Value
    '     文字の種類（内容はgetCharList()の引数（alType）と同じ）
    '---------------------------------------------------------------------------------------------------
    'History
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
            this_createDefinitionsByCharacterType lPowerOf2
            If PoChar2Type.Exists(bCode) Then
                whatType = PoChar2Type.Item(bCode)
                Exit Function
            End If
            lPowerOf2 = lPowerOf2+1
        Loop
    End Function
    
    '***************************************************************************************************
    'Function/Sub Name           : charList()
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
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/28         Y.Fujii                  First edition
    '***************************************************************************************************
    Public Function charList( _
        byVal alType _
        )
        Dim lType : lType = alType
        Dim lPowerOf2 : lPowerOf2 = Cl_MAX_POWER_OF_2
        Dim vRet,lTargetType
        Do Until lPowerOf2<0
            lTargetType = 2^lPowerOf2
            If (lType-lTargetType)>=0 Then
                lType = lType-lTargetType
                this_createDefinitionsByCharacterType lPowerOf2
                cf_pushA vRet, PoType2Chars.Item(lTargetType)
            End If
            lPowerOf2 = lPowerOf2 - 1
        Loop
        charList = vRet
    End Function
    

    '***************************************************************************************************
    'Function/Sub Name           : this_createDefinitionsByCharacterType
    'Overview                    : 指定した文字種類の定義を作成する
    'Detailed Description        : 工事中
    'Argument
    '     alPowerOf2             : 文字の種類（内容はgetCharList()の引数（alType）と同じ）を2^nとした場合のn
    'Return Value
    '     なし
    '---------------------------------------------------------------------------------------------------
    'History
    'Date               Name                     Reason for Changes
    '----------         ----------------------   -------------------------------------------------------
    '2023/10/30         Y.Fujii                  First edition
    '***************************************************************************************************
    Private Sub this_createDefinitionsByCharacterType( _
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
