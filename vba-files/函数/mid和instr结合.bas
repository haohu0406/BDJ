mid和instr结合取得 《》和 ”“中间的内容
For l = 1 To UBound(arr_col, 2)
    smhl = InStr(1, arr_col(1, l), "《")
    smhr = InStr(1, arr_col(1, l), "》")
    syhl = InStr(smhr, arr_col(1, l), "“")
    syhr = InStr(smhr, arr_col(1, l), "”")
    arr_sh(l) = Mid(arr_col(1, l), smhl + 1, smhr - smhl - 1)
    arr_rng(l) = Mid(arr_col(1, l), syhl + 1, syhr - syhl - 1)
Next
