Attribute VB_Name = "ModNumeroATexto"
Function NUMEROATEXTO(Numero As Double) As String

Dim Letra As String
Const Maximo = 1999999999.99

If (Numero >= 0) And (Numero <= Maximo) Then
    Letra = NUMERORECURSIVO((Fix(Numero)))
    NUMEROATEXTO = Letra
Else 
    NUMEROATEXTO = "ERROR: El numero excede los limites."
End If

End Function

Function NUMERORECURSIVO(Numero As Long) As String

Dim Unidades, Decenas, Centenas
Dim Resultado As String

'**************************************************
' Nombre de los numeros
'**************************************************
Unidades = Array("", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciseis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiuno", "veintidos", "veintitres", "veinticuatro", "veinticinco", "veintiseis", "veintisiete", "veintiocho", "veintinueve")
Decenas = Array("", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa", "cien")
Centenas = Array("", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos")
'**************************************************

Select Case Numero
    Case 0
        Resultado = "cero"
    Case 1 To 29
        Resultado = Unidades(Numero)
    Case 30 To 100
        Resultado = Decenas(Numero \ 10) + IIf(Numero Mod 10 <> 0, " y " + NUMERORECURSIVO(Numero Mod 10), "")
    Case 101 To 999
        Resultado = Centenas(Numero \ 100) + IIf(Numero Mod 100 <> 0, " " + NUMERORECURSIVO(Numero Mod 100), "")
    Case 1000 To 1999
        Resultado = "mil" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 2000 To 999999
        Resultado = NUMERORECURSIVO(Numero \ 1000) + " mil" + IIf(Numero Mod 1000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000), "")
    Case 1000000 To 1999999
        Resultado = "un millon" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
    Case 2000000 To 1999999999
        Resultado = NUMERORECURSIVO(Numero \ 1000000) + " millones" + IIf(Numero Mod 1000000 <> 0, " " + NUMERORECURSIVO(Numero Mod 1000000), "")
End Select

NUMERORECURSIVO = Resultado

End Function
