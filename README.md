'PAR DOUGHMI NAWFAL
'INGENIEUR D'ETAT EN GENIE INDUSTRIEL OPTION LOGISTIQUE INTERNATIONALE
'ESITH CASABLANCA
Sub Post_fix()
Dim i, j As Integer
Dim a, ar, b, br, c, cr As Integer
a = 0
ar = 0
'LA PREMIERE BOUCLE POUR LES COMPAGNIE CLIENTES
'LA DEUXIEME BOUCLE EST POUR LA COMPAGNIE RAM
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 2), Worksheets("pm").Cells(i, 7)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 2), Worksheets("pm").Cells(j, 7)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 2) = a + ar
b = 0
br = 0
'***************06H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 8), Worksheets("pm").Cells(i, 13)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
b = b + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 8), Worksheets("pm").Cells(j, 13)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
br = br + 1
End If
Next j
Worksheets("pm").Cells(305, 8) = b + br
c = 0
cr = 0
'***************07H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 14), Worksheets("pm").Cells(i, 19)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
c = c + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 14), Worksheets("pm").Cells(j, 19)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
cr = cr + 1
End If
Next j
Worksheets("pm").Cells(305, 14) = c + cr
a = 0
ar = 0
'***************08H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 20), Worksheets("pm").Cells(i, 25)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 20), Worksheets("pm").Cells(j, 25)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 20) = a + ar
a = 0
ar = 0
'***************09H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 26), Worksheets("pm").Cells(i, 31)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 26), Worksheets("pm").Cells(j, 31)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 26) = a + ar
a = 0
ar = 0
'***************10H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 32), Worksheets("pm").Cells(i, 37)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 32), Worksheets("pm").Cells(j, 37)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 32) = a + ar
a = 0
ar = 0
'***************11H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 38), Worksheets("pm").Cells(i, 43)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 38), Worksheets("pm").Cells(j, 43)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 38) = a + ar
a = 0
ar = 0
'***************12H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 44), Worksheets("pm").Cells(i, 49)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 44), Worksheets("pm").Cells(j, 49)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 44) = a + ar
a = 0
ar = 0
'***************13H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 50), Worksheets("pm").Cells(i, 55)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 50), Worksheets("pm").Cells(j, 55)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 50) = a + ar
a = 0
ar = 0
'***************14H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 56), Worksheets("pm").Cells(i, 61)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 56), Worksheets("pm").Cells(j, 61)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 56) = a + ar
a = 0
ar = 0
'***************15H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 62), Worksheets("pm").Cells(i, 67)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 62), Worksheets("pm").Cells(j, 67)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 62) = a + ar
a = 0
ar = 0
'***************16H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 68), Worksheets("pm").Cells(i, 73)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 68), Worksheets("pm").Cells(j, 73)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 68) = a + ar
a = 0
ar = 0
'***************17H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 74), Worksheets("pm").Cells(i, 79)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 74), Worksheets("pm").Cells(j, 79)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 74) = a + ar
a = 0
ar = 0
'***************18H00****************************
For i = 21 To 222
If Range(Worksheets("pm").Cells(i, 80), Worksheets("pm").Cells(i, 85)).Interior.Pattern = xlNone And Worksheets("pm").Cells(i, 1) <> "" Then
a = a + 1
End If
Next i
For j = 253 To 294
If Range(Worksheets("pm").Cells(j, 80), Worksheets("pm").Cells(j, 85)).Interior.Pattern = xlNone And Worksheets("pm").Cells(j, 1) <> "" Then
ar = ar + 1
End If
Next j
Worksheets("pm").Cells(305, 80) = a + ar


Worksheets("pm").Cells(305, 1).Select

End Sub
