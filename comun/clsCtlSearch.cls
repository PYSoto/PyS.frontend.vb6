VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub formSearch(objeto As Variant, KeyAscii As Integer, cadena As String, db As clsDB)

    Load frmSearch
    
    With frmSearch
        .objeto = objeto
        .db = db
        .txtCadena.Text = Chr(KeyAscii)
        .Caption = "Búsqueda " & cadena
        .Show vbModal
    End With
    
    Unload frmSearch
    
End Sub


