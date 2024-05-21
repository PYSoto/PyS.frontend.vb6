VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEncontrados 
      Height          =   5715
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtCadena 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vObjeto As Variant

Private vDB As clsDB

Public Property Get objeto() As Variant

    Set objeto = vObjeto
    
End Property

Public Property Let objeto(vNewValue As Variant)

    Set vObjeto = vNewValue
    
End Property

Public Property Get db() As Variant

    Set db = vDB
    
End Property

Public Property Let db(vNewValue As Variant)

    Set vDB = vNewValue
    
End Property

Private Sub cmdSalir_Click()

    Me.Hide
    
End Sub

Private Sub Form_Activate()

    Me.txtCadena.SelStart = Len(Me.txtCadena.Text)
    
End Sub

Private Sub Form_Load()

    Set vDB = Nothing
    
End Sub

Private Sub lstEncontrados_DblClick()

    If Me.lstEncontrados.ListIndex < 0 Then Exit Sub
    
    vObjeto.findSearch Me.lstEncontrados.ItemData(Me.lstEncontrados.ListIndex), vDB
    
    cmdSalir_Click

End Sub

Private Sub lstEncontrados_KeyPress(keyAscii As Integer)

    If Me.lstEncontrados.ListIndex < 0 Then Exit Sub
    If keyAscii <> 13 Then Exit Sub
    
    vObjeto.findSearch Me.lstEncontrados.ItemData(Me.lstEncontrados.ListIndex), vDB
    
    cmdSalir_Click

End Sub

Private Sub txtCadena_Change()
Dim varObjeto As Variant

Dim objetos As Collection

    Set objetos = vObjeto.collectionSearch(Me.txtCadena.Text, vDB)
    
    Me.lstEncontrados.Clear
    
    For Each varObjeto In objetos
        Me.lstEncontrados.AddItem varObjeto.textSearch
        Me.lstEncontrados.ItemData(Me.lstEncontrados.NewIndex) = varObjeto.idSearch
    Next
    
End Sub
