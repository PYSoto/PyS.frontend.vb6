VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificación Usuario"
   ClientHeight    =   1230
   ClientLeft      =   2820
   ClientTop       =   2835
   ClientWidth     =   5190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "&Ingresar"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtContrasenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblContrasenha 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   810
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIngresar_Click()
Dim objUA As New clsDAOUsuarioActivo
    
    If Not gUsuario.usuarioValido(DB) Then
        MsgBox "ERROR: Usuario o Contraseña INCORRECTA"
        txtUsuario.Text = ""
        txtUsuario.SetFocus
    Else
        objUA.grabaUsuario gMDI, gUsuario.nombre, DB
        Unload Me
    End If
    
End Sub

Private Sub cmdSalir_Click()

    End

End Sub

Private Sub txtContrasenha_GotFocus()

    marcarseleccion Me.txtContrasenha

End Sub

Private Sub txtContrasenha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdIngresar.SetFocus
    End If
End Sub

Private Sub txtContrasenha_LostFocus()

    gUsuario.password = Me.txtContrasenha.Text

End Sub

Private Sub txtUsuario_GotFocus()

    marcarseleccion Me.txtUsuario

End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.txtContrasenha.SetFocus
    End If

End Sub
Private Sub txtUsuario_LostFocus()

    gUsuario.nombre = Me.txtUsuario.Text

End Sub
