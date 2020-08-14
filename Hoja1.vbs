Private Sub btIdentificar_Click()
    Dim excepInicial() As String
    Dim excepTotal() As String

    Dim i As Integer
    Dim address2Last As Variant
    
    address2Last = Split(Hoja1.Range("FORMULAS").Address, "$")
    Hoja1.Range("FORMULAS").AutoFill Destination:=Hoja1.Range(address2Last(1) & address2Last(2) & address2Last(3) & Hoja1.Range("A1").End(xlDown).Row)
    
    i = 0
    While Hoja3.Range("EXCLUSION_INICIAL").Offset(i, 0) <> ""
        ReDim Preserve excepInicial(i)
        excepInicial(i) = Hoja3.Range("EXCLUSION_INICIAL").Offset(i, 0)
        i = i + 1
    Wend
    
    i = 0
    While Hoja3.Range("EXCLUSION_ESPECIFICA").Offset(i, 0) <> ""
        ReDim Preserve excepTotal(i)
        excepTotal(i) = Hoja3.Range("EXCLUSION_ESPECIFICA").Offset(i, 0)
        i = i + 1
    Wend
    
    i = 1
    Do While Hoja1.Range("GLOSA").Offset(i, 0).Value <> ""
        If IsInArray(Hoja1.Range("GLOSA").Offset(i, 0).Value, excepInicial) Then
            Hoja1.Range("IDENTIFICABLE").Offset(i, 0) = "FALSO"
        Else
            If IsInArrayExact(Hoja1.Range("GLOSA").Offset(i, 0).Value, excepTotal) Then
                Hoja1.Range("IDENTIFICABLE").Offset(i, 0) = "FALSO"
            Else
                Hoja1.Range("IDENTIFICABLE").Offset(i, 0) = "VERDADERO"
            End If
        End If
        i = i + 1
    Loop
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim sw As Boolean
    sw = False
    For Each estr In arr
        If InStr(1, stringToBeFound, estr, vbTextCompare) Then
            sw = True
            Exit For
        End If
    Next estr
    IsInArray = sw
End Function


Function IsInArrayExact(stringToBeFound As String, arr As Variant) As Boolean
    Dim sw As Boolean
    sw = False
    For Each estr In arr
        If stringToBeFound = estr Then
            sw = True
            Exit For
        End If
    Next estr
    IsInArrayExact = sw
End Function

Private Sub CommandButton1_Click()
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
                       TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                       Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
                       FieldInfo:=Array(1, 4), TrailingMinusNumbers:=True
    Selection.NumberFormat = "YYYY-MM-DD"
End Sub
