VERSION 5.00
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEncontrados 
      Height          =   2010
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
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vObjeto As Variant

Private vColeccion As Collection

Private Sub cmdSalir_Click()

    Me.Hide
    
End Sub

Public Sub setObjeto(vNewValue As Variant)

    Set vObjeto = vNewValue
    
End Sub

Private Sub Form_Activate()

    Me.txtCadena.SetFocus
    Me.txtCadena.SelStart = Len(Me.txtCadena.Text)
    
End Sub

Private Sub lstEncontrados_Click()

    If Me.lstEncontrados.ListIndex < 0 Then Exit Sub
    
    vObjeto.cargarBuscar Me.lstEncontrados.ItemData(Me.lstEncontrados.ListIndex), db
    
    cmdSalir_Click
    
End Sub

Private Sub txtCadena_Change()
Dim varObjeto As Variant

    Set vColeccion = vObjeto.coleccionBuscar(Me.txtCadena.Text, db)
    
    Me.lstEncontrados.Clear
    
    For Each varObjeto In vColeccion
        Me.lstEncontrados.AddItem varObjeto.descripcionBuscar
        Me.lstEncontrados.ItemData(Me.lstEncontrados.NewIndex) = varObjeto.codigoBuscar
    Next
    
End Sub
