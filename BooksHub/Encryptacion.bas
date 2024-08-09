Attribute VB_Name = "Encryptacion"
Public Function Encrypt(ByVal Word As String, ByVal Key As String, _
Optional ByVal Mode As Boolean = False) As String
    Dim w As Long, k As Long, p As Long, j As Long, NuChr As Long
    Dim Cd As String, Kd As String, Rd As String
    w = Len(Word)
    k = Len(Key)
    ' Modalidad de Encritacion...
    If Mode = False Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) + Asc(Kd)
            If NuChr > 255 Then
                NuChr = NuChr - 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function
    End If
    ' Modalidad de Desencriptacion...
    If Mode = True Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) - Asc(Kd)
            If NuChr < 0 Then
                NuChr = NuChr + 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function
    End If
End Function

