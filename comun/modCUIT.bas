Attribute VB_Name = "modCUIT"
Option Explicit

Public Function validaCUIT(ByVal pCUIT As String) As Boolean

Dim intValor As Integer
Dim intVerificador As Integer

    pCUIT = Trim(pCUIT)
    
    If Len(pCUIT) <> 11 Then
        validaCUIT = False
        Exit Function
    End If
    
    If pCUIT = String(11, "0") Or pCUIT = String(11, "1") Then
        validaCUIT = False
        Exit Function
    End If
    
    intValor = 5 * CInt(Mid(pCUIT, 1, 1))
    intValor = intValor + 4 * CInt(Mid(pCUIT, 2, 1))
    intValor = intValor + 3 * CInt(Mid(pCUIT, 3, 1))
    intValor = intValor + 2 * CInt(Mid(pCUIT, 4, 1))
    intValor = intValor + 7 * CInt(Mid(pCUIT, 5, 1))
    intValor = intValor + 6 * CInt(Mid(pCUIT, 6, 1))
    intValor = intValor + 5 * CInt(Mid(pCUIT, 7, 1))
    intValor = intValor + 4 * CInt(Mid(pCUIT, 8, 1))
    intValor = intValor + 3 * CInt(Mid(pCUIT, 9, 1))
    intValor = intValor + 2 * CInt(Mid(pCUIT, 10, 1))
    
    intValor = intValor Mod 11
    
    intVerificador = 11 - intValor
    
    Select Case intVerificador
        Case 11
            intVerificador = 0
        Case 10
            intVerificador = 9
    End Select
    
    validaCUIT = IIf(intVerificador = CInt(Mid(pCUIT, 11, 1)), True, False)
    
End Function

Public Function generaCUIT(ByVal pDNI As String, ByVal pSexo As String) As String
Dim strXY As String

Dim intValor As Integer
Dim intVerificador As Integer

    pDNI = Right("00000000" & Trim(pDNI), 8)

    pSexo = UCase(pSexo)
    
    Select Case pSexo
        Case "M"
            strXY = "20"
        Case "F"
            strXY = "27"
        Case Else
            strXY = "30"
    End Select
    
    intValor = 5 * CInt(Mid(strXY, 1, 1))
    intValor = intValor + 4 * CInt(Mid(strXY, 2, 1))
    intValor = intValor + 3 * CInt(Mid(pDNI, 1, 1))
    intValor = intValor + 2 * CInt(Mid(pDNI, 2, 1))
    intValor = intValor + 7 * CInt(Mid(pDNI, 3, 1))
    intValor = intValor + 6 * CInt(Mid(pDNI, 4, 1))
    intValor = intValor + 5 * CInt(Mid(pDNI, 5, 1))
    intValor = intValor + 4 * CInt(Mid(pDNI, 6, 1))
    intValor = intValor + 3 * CInt(Mid(pDNI, 7, 1))
    intValor = intValor + 2 * CInt(Mid(pDNI, 8, 1))
    
    intValor = intValor Mod 11
    
    intVerificador = 11 - intValor
    
    If intValor = 0 Then intVerificador = 0
    
    If intValor = 1 Then
        Select Case pSexo
            Case "M"
                intVerificador = 9
                strXY = "23"
            Case "F"
                intVerificador = 4
                strXY = "23"
        End Select
    End If
    
    generaCUIT = strXY & "-" & pDNI & "-" & Trim(Str(intVerificador))
    
End Function
