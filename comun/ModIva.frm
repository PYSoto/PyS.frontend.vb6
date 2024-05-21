VERSION 5.00
Begin VB.Form ModifIva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alícuotas Débito Fiscal IVA"
   ClientHeight    =   1350
   ClientLeft      =   2745
   ClientTop       =   1620
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4110
   Begin VB.TextBox txtNoInscripto 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtInscripto 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Resp No Inscripto"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Resp Inscripto"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "ModifIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objPar As New clsDAOParametros

Private Sub cmdGuardar_Click()

    With objPar
        .findLast
        
        .iva1 = Val(Me.txtInscripto.Text)
        .iva2 = Val(Me.txtNoInscripto.Text)
        
        .save
    End With
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    objPar.findLast

    Me.txtInscripto.Text = Format(objPar.iva1, "0.00")
    Me.txtNoInscripto.Text = Format(objPar.iva2, "0.00")
    
End Sub

Private Sub txtInscripto_GotFocus()

    marcarseleccion Me.txtInscripto
    
End Sub

Private Sub txtNoInscripto_GotFocus()

    marcarseleccion Me.txtNoInscripto
    
End Sub
