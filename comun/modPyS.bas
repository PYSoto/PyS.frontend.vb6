Attribute VB_Name = "modPyS"
Option Explicit

Global db As New clsDB
Global au As New clsDB

Global gCon As New clsCtlConnect

Global gEmpresa As New clsDAOEmpresa

Global gUsuario As New clsDAOUsuarios

Global gMDI As MDIForm

Global gArtMirror As New clsCtlMirror

Global proveedorglobal As New clsDAOProveedor

Global Const cntArchivo = "P&S"

Global Const cntProhibidos = "/[^<,.@\{}()*$%?=>:|;#]+'"""

Global Const cntTTCompras = 2
Global Const cntTTPagos = 4

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function evaluarString(ByVal pCadena As String) As String
Dim strCadena As String
Dim intCiclo As Integer
    
    strCadena = ""
    For intCiclo = 1 To Len(pCadena)
        strCadena = strCadena & Mid(pCadena, intCiclo, 1)
        If Mid(pCadena, intCiclo, 1) = "\" Then strCadena = strCadena & "\"
        If Mid(pCadena, intCiclo, 1) = "'" Then strCadena = strCadena & "'"
        If Mid(pCadena, intCiclo, 1) = Chr(34) Then strCadena = strCadena & Chr(34)
    Next intCiclo
    
    evaluarString = strCadena
    
End Function

Public Sub mesSiguiente(ByRef pMes As Integer, ByRef pAnho As Integer)
    
    pMes = pMes + 1
    If pMes = 13 Then
        pMes = 1
        pAnho = pAnho + 1
    End If

End Sub

Public Sub mesAnterior(ByRef pMes As Integer, ByVal pAnho As Integer)
    
    pMes = pMes - 1
    If pMes = 0 Then
        pMes = 12
        pAnho = pAnho - 1
    End If

End Sub

Public Sub marcarseleccion(ByRef pTextBox As TextBox)
    
    pTextBox.SelStart = 0
    pTextBox.SelLength = Len(pTextBox.Text)

End Sub

Public Sub llenaComboIVA(ByRef pCombo As ComboBox)
Dim strPosIVA As Variant
Dim intCodIVA As Variant

Dim intCiclo As Integer

    strPosIVA = Array("Resp.Inscripto", "Consum.Final", "Monotributista", "Resp.No Inscripto", "Exento", "Exportación")
    intCodIVA = Array(1, 2, 3, 4, 5, 6)

    pCombo.Clear
    For intCiclo = LBound(strPosIVA) To UBound(strPosIVA)
        pCombo.AddItem strPosIVA(intCiclo)
        pCombo.ItemData(pCombo.NewIndex) = intCodIVA(intCiclo)
    Next intCiclo
    
    If pCombo.ListCount > 0 Then pCombo.Text = strPosIVA(LBound(strPosIVA))

End Sub

Public Function a2Decimales(ByVal pValor As Double) As Double
    
    a2Decimales = CDbl(CLng(pValor * 100)) / 100

End Function

Public Function cambiarTabPorSpace(ByVal pCadena As String) As String
Dim strCadena As String

Dim intContador As Integer

    strCadena = ""
    
    For intContador = 1 To Len(pCadena)
        If Asc(Mid(pCadena, intContador, 1)) = 9 Then
            strCadena = strCadena & Chr(32)
        Else
            strCadena = strCadena & Mid(pCadena, intContador, 1)
        End If
    Next intContador
    
    cambiarTabPorSpace = strCadena
    
End Function

Public Function tieneTab(ByVal pCadena As String) As Boolean
Dim intContador As Integer

    tieneTab = False
    
    For intContador = 1 To Len(pCadena)
        If Asc(Mid(pCadena, intContador, 1)) = 9 Then
            tieneTab = True
            Exit Function
        End If
    Next intContador

End Function

Public Function entreComillas(ByVal pCadena As Variant) As String
    
    entreComillas = Chr(34) & pCadena & Chr(34)
    
End Function

Public Function sacarComillas(ByVal pCadena As String) As String

    sacarComillas = pCadena
    
    If Left(pCadena, 1) = Chr(34) And Right(pCadena, 1) = Chr(34) Then sacarComillas = Mid(pCadena, 2, Len(pCadena) - 2)
    
End Function

Public Function getPath(pArchivo As String, pFileName As String) As String
Dim intPosicion As Integer

    intPosicion = InStrRev(pArchivo, "\")
    getPath = Mid(pArchivo, 1, intPosicion)
    pFileName = Mid(pArchivo, intPosicion + 1)
    
End Function

Public Sub fillComboIva(combo As ComboBox)
Dim posicionIva As Variant
Dim codigoIva As Variant

Dim ciclo As Integer

    posicionIva = Array("Resp.Inscripto", "Consum.Final", "Monotributista", "Resp.No Inscripto", "Exento", "Exportación")
    codigoIva = Array(1, 2, 3, 4, 5, 6)

    combo.Clear
    For ciclo = LBound(posicionIva) To UBound(posicionIva)
        combo.AddItem posicionIva(ciclo)
        combo.ItemData(combo.NewIndex) = codigoIva(ciclo)
    Next ciclo
    
    If combo.ListCount > 0 Then combo.Text = posicionIva(LBound(posicionIva))

End Sub


