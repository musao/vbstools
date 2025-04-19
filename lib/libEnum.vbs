'***************************************************************************************************
'FILENAME                    : libEnum.vbs
'Overview                    : Enum
'Detailed Description        : çHéñíÜ
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2025/03/06         Y.Fujii                  First edition
'***************************************************************************************************

Call new_Enum( _
    "topic" _
    , new_DicOf( _
        Array( _
            "LOG", 1 _
        ) _
    ) _
)
Call new_Enum( _
    "logType" _
    , new_DicOf( _
        Array( _
            "ERROR", 1 _
            , "WARNING", 3 _
            , "INFO", 5 _
            , "DEBUG", 7 _
            , "TRACE", 9 _
        ) _
    ) _
)
'            , "DETAIL", 11 _
Call new_Enum( _
    "charType" _
    , new_DicOf( _
        Array( _
            "HALF_WIDTH_ALPHABET_UPPERCASE", 2^0 _
            , "HALF_WIDTH_ALPHABET_LOWERCASE", 2^1 _
            , "HALF_WIDTH_NUMBERS", 2^2 _
            , "HALF_WIDTH_SYMBOL", 2^3 _
            , "HALF_WIDTH_KATAKANA", 2^4 _
            , "HALF_WIDTH_KATAKANA_SYMBOL", 2^5 _
            , "FULL_WIDTH_ALPHABET_UPPERCASE", 2^6 _
            , "FULL_WIDTH_ALPHABET_LOWERCASE", 2^7 _
            , "FULL_WIDTH_NUMBERS", 2^8 _
            , "FULL_WIDTH_SYMBOL", 2^9 _
            , "FULL_WIDTH_HIRAGANA", 2^10 _
            , "FULL_WIDTH_KATAKANA", 2^11 _
            , "FULL_WIDTH_GREEK_CYRILLIC_UPPERCASE", 2^12 _
            , "FULL_WIDTH_GREEK_CYRILLIC_LOWERCASE", 2^13 _
            , "FULL_WIDTH_LINEFRAME", 2^14 _
            , "FULL_WIDTH_KANJI_LEVEL1", 2^15 _
            , "FULL_WIDTH_KANJI_LEVEL2", 2^16 _
        ) _
    ) _
)
