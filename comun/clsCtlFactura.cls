VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub print_factura(clientemovimientoId As Long, tipocomprobante As String, puntoventa As Integer, numerocomprobante As Long, hWnd As Long)
Dim request As MSXML2.ServerXMLHTTP

Dim file_stream As ADODB.Stream

Dim url As String
Dim filename As String

    filename = Environ$("tmp") & "\comprobante." & tipocomprobante & Format(puntoventa, "0000") & "-" & Format(numerocomprobante, "00000000") & ".pdf"

    Set request = New MSXML2.ServerXMLHTTP
    
    url = modUrls.url_pys() & "makefactura/pdf/" & clientemovimientoId

    request.setTimeouts 200000, 200000, 200000, 200000
    
    request.Open "GET", url
    request.send
    
    If request.Status = 200 Then
        Set file_stream = New ADODB.Stream
        file_stream.Open
        file_stream.Type = adTypeBinary
        
        file_stream.Write request.responseBody
        file_stream.Position = 0
        
        file_stream.SaveToFile filename, adSaveCreateOverWrite
        file_stream.Close
        
        Set file_stream = Nothing
        
        ShellExecute hWnd, "open", filename, vbNullString, vbNullString, 1
        
    End If
    
    Set request = Nothing
    
End Sub

