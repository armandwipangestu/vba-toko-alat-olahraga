Attribute VB_Name = "ModuleBulan"
Option Explicit

Public Function convertBulanIndonesia(tanggal As Date) As String
    Dim bulan As String
    bulan = Format(tanggal, "mm")
    
    If bulan = "01" Then
        convertBulanIndonesia = "Januari"
    End If
    
    If bulan = "02" Then
        convertBulanIndonesia = "Februari"
    End If
    
    If bulan = "03" Then
        convertBulanIndonesia = "Maret"
    End If
    
    If bulan = "04" Then
        convertBulanIndonesia = "April"
    End If
    
    If bulan = "05" Then
        convertBulanIndonesia = "Mei"
    End If
    
    If bulan = "06" Then
        convertBulanIndonesia = "Juni"
    End If
    
    If bulan = "07" Then
        convertBulanIndonesia = "Juli"
    End If
    
    If bulan = "08" Then
        convertBulanIndonesia = "Agustus"
    End If
    
    If bulan = "09" Then
        convertBulanIndonesia = "September"
    End If
    
    If bulan = "10" Then
        convertBulanIndonesia = "Oktober"
    End If
    
    If bulan = "11" Then
        convertBulanIndonesia = "November"
    End If
    
    If bulan = "12" Then
        convertBulanIndonesia = "Desember"
    End If
    
End Function

Public Function convertHariIndonesia(hari As String) As String
    If hari = "Sunday" Then
        convertHariIndonesia = "Minggu"
    End If
    
    If hari = "Monday" Then
        convertHariIndonesia = "Senin"
    End If
    
    If hari = "Tuesday" Then
        convertHariIndonesia = "Selasa"
    End If
    
    If hari = "Wednesday" Then
        convertHariIndonesia = "Rabu"
    End If
    
    If hari = "Thursday" Then
        convertHariIndonesia = "Kamis"
    End If
    
    If hari = "Friday" Then
        convertHariIndonesia = "Jum'at"
    End If
    
    If hari = "Saturday" Then
        convertHariIndonesia = "Sabtu"
    End If
End Function

