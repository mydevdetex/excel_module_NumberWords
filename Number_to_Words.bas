Attribute VB_Name = "Number_to_Words"
Function NumberWords(angka As Variant) As String
'untuk angka satuan
    If angka < 12 Then
        NumberWords = GetSatuan(angka)
'untuk angka belasan
    ElseIf angka < 20 Then
        NumberWords = GetBelasan(angka)
'untuk angka puluhan
    ElseIf angka < 100 Then
        NumberWords = GetPuluhan(angka)
'untuk angka ratusan
    ElseIf angka < 1000 Then
        NumberWords = GetRatusan(angka)
'untuk angka ribuan
    ElseIf angka < 10000 Then
        NumberWords = GetRibuan(angka)
'untuk angka belas ribuan sampai ratus ribuan
    ElseIf angka < 1000000 Then
        NumberWords = GetRatusRibuan(angka)
'untuk angka jutaan
    ElseIf angka < 1000000000 Then
        NumberWords = GetJutaan(angka)
'untuk angka miliaran
    ElseIf angka < 1000000000000# Then
        NumberWords = GetMiliaran(angka)
'untuk angka triliunan
    Else
        NumberWords = "sorry it's out of range"
    End If
End Function
Private Function ConvertSubNumbers(angka As Variant) As String
    Dim satuan As Variant
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka < 12 Then
    'satuan
        ConvertSubNumbers = satuan(angka)
    ElseIf angka < 20 Then
    'belasan
        ConvertSubNumbers = GetBelasan(angka)
    ElseIf angka < 100 Then
    'puluhan
        ConvertSubNumbers = GetPuluhan(angka)
    ElseIf angka < 1000 Then
    'ratusan
        ConvertSubNumbers = GetRatusan(angka)
    ElseIf angka < 12000 Then
    'ribuan
        ConvertSubNumbers = GetRibuan(angka)
    ElseIf angka < 1000000 Then
    'belasan ribu sampai ratusan ribu
        ConvertSubNumbers = GetRatusRibuan(angka)
    'jutaan
    ElseIf angka < 1000000000 Then
        ConvertSubNumbers = GetJutaan(angka)
    'miliar
    Else
        ConvertSubNumbers = GetMiliaran(angka)
    End If
End Function
Private Function GetSatuan(angka As Variant) As String
    Dim satuan As Variant
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka = 0 Then
            GetSatuan = "nol"
        Else
            GetSatuan = satuan(angka)
        End If
End Function
Private Function GetBelasan(angka As Variant) As String
    Dim satuan As Variant
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka < 1 Then
        GetBelasan = ""
    Else
        GetBelasan = satuan(Mid(CStr(angka), 2, 1)) & " belas"
    End If
End Function
Private Function GetPuluhan(angka As Variant) As String
    Dim satuan As Variant
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka < 1 Then
        GetPuluhan = ""
    Else
        GetPuluhan = satuan(Mid(CStr(angka), 1, 1)) & " puluh " & satuan(Mid(CStr(angka), 2, 1))
    End If
End Function
Private Function GetRatusan(angka As Variant) As String
    Dim satuan As Variant, se As String
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka < 1 Then
        GetRatusan = ""
    End If
    If Mid(CStr(angka), 1, 1) = 1 Then
        se = "seratus"
    Else
        se = satuan(Mid(CStr(angka), 1, 1)) & " ratus"
    End If
    GetRatusan = se & " " & ConvertSubNumbers(Mid(CStr(angka), 2, 2))
End Function
Private Function GetRibuan(angka As Variant) As String
    Dim satuan As Variant, se As String, ratusan As String
    satuan = Array("", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
    If angka < 1 Then
        GetRibuan = ""
    End If
    If Mid(CStr(angka), 1, 1) = 1 Then
        se = "seribu"
    Else
        se = satuan(Mid(CStr(angka), 1, 1)) & " ribu"
    End If
    GetRibuan = se & " " & ConvertSubNumbers(Mid(CStr(angka), 2, 3))
End Function
Private Function GetRatusRibuan(angka As Variant) As String
    Dim ratusribu As String, ratus As String, ribu As String
    ribu = " ribu"
    If angka < 1 Then
        GetRatusRibuan = ""
    Else
        If Len(CStr(angka)) = 5 Then
            ratus = ConvertSubNumbers(Mid(CStr(angka), 3, 3))
            ratusribu = ConvertSubNumbers(Mid(CStr(angka), 1, 2))
        ElseIf Len(CStr(angka)) = 6 Then
            ratus = ConvertSubNumbers(Mid(CStr(angka), 4, 3))
            ratusribu = ConvertSubNumbers(Mid(CStr(angka), 1, 3))
        End If
        GetRatusRibuan = ratusribu & ribu & " " & ratus
    End If
End Function
Private Function GetJutaan(angka As Variant) As String
    Dim ratusribu As String, juta As String
    If angka < 1 Then
        GetJutaan = ""
    Else
        If Len(CStr(angka)) = 7 Then
            ratusribu = ConvertSubNumbers(Mid(CStr(angka), 2, 6))
            juta = ConvertSubNumbers(Mid(CStr(angka), 1, 1))
        ElseIf Len(CStr(angka)) = 8 Then
            ratusribu = ConvertSubNumbers(Mid(CStr(angka), 3, 6))
            juta = ConvertSubNumbers(Mid(CStr(angka), 1, 2))
        ElseIf Len(CStr(angka)) = 9 Then
            ratusribu = ConvertSubNumbers(Mid(CStr(angka), 4, 6))
            juta = ConvertSubNumbers(Mid(CStr(angka), 1, 3))
        End If
        GetJutaan = juta & " juta" & " " & ratusribu
    End If
End Function
Private Function GetMiliaran(angka As Variant) As String
    Dim juta As String, miliar As String
    If angka < 1 Then
        GetMiliaran = ""
    Else
        If Len(CStr(angka)) = 10 Then
            juta = ConvertSubNumbers(Mid(CStr(angka), 2, 9))
            miliar = ConvertSubNumbers(Mid(CStr(angka), 1, 1))
        ElseIf Len(CStr(angka)) = 11 Then
            juta = ConvertSubNumbers(Mid(CStr(angka), 3, 9))
            miliar = ConvertSubNumbers(Mid(CStr(angka), 1, 2))
        ElseIf Len(CStr(angka)) = 12 Then
            juta = ConvertSubNumbers(Mid(CStr(angka), 4, 9))
            miliar = ConvertSubNumbers(Mid(CStr(angka), 1, 3))
        End If
        GetMiliaran = miliar & " miliar" & " " & juta
    End If
End Function
