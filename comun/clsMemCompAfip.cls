VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemCompAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCodigo As Integer
Private vDescripcion As String

Private vCollection As New Collection

Private Sub Class_Initialize()

    vCodigo = 0
    vDescripcion = ""
    
    fillCollection
    
End Sub

Public Property Get codigo() As Variant

    codigo = vCodigo

End Property

Public Property Let codigo(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigo = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Private Sub fillCollection()

    vCollection.add createItem(1, "Facturas A")
    vCollection.add createItem(3, "Notas de Credito A")
    vCollection.add createItem(6, "Facturas B")
    vCollection.add createItem(8, "Notas de Credito B")

End Sub

Private Function createItem(pCodigo As Integer, pDescripcion As String) As clsMemCompAfip
Dim objCA As New clsMemCompAfip

    With objCA
        .codigo = pCodigo
        .descripcion = pDescripcion
    End With
    
    Set createItem = objCA
    
End Function
