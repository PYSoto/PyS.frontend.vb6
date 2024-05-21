VERSION 5.00
Begin VB.Form frmCambiarClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Clave"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5985
   Begin VB.TextBox txtClaveVieja 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtReClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdCambiarClave 
      Caption         =   "&Cambiar Clave"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblUsuario 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Clave Anterior"
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   990
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Re - Clave"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   750
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.Label etiq 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   405
   End
End
Attribute VB_Name = "frmCambiarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtClave_GotFocus()
    
    marcarseleccion Me.txtClave

End Sub

Private Sub cmdCambiarClave_Click()
Dim strSQL As String

Dim rstQuery As Recordset
    
    If Trim(txtClaveVieja.Text) = "" Then
        MsgBox "Falta CLAVE . . ."
        txtClaveVieja.SetFocus
        Exit Sub
    End If
    If Trim(txtClave.Text) = "" Then
        MsgBox "Falta CLAVE . . ."
        txtClave.SetFocus
        Exit Sub
    End If
    If txtClave.Text <> txtReClave.Text Then
        MsgBox "Claves NO Coinciden . . ."
        txtReClave.SetFocus
        Exit Sub
    End If
    
    If Not gUsuario.usuarioValido(db) Then
        MsgBox "Contraseña INCORRECTA"
        Me.txtClaveVieja.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    With gUsuario
        .password = Me.txtClave.Text
        .update
    End With
    
    txtClaveVieja.Text = ""
    txtClave.Text = ""
    txtReClave.Text = ""
    Me.MousePointer = 0

End Sub

Private Sub Form_Activate()
    
    Me.lblUsuario.Caption = gUsuario.nombre
    
    txtClaveVieja.Text = ""
    txtClave.Text = ""
    txtReClave.Text = ""

End Sub

Private Sub txtClaveVieja_Change()

    gUsuario.password = Me.txtClaveVieja.Text

End Sub

Private Sub txtClaveVieja_GotFocus()

    marcarseleccion Me.txtClaveVieja

End Sub

Private Sub txtReClave_GotFocus()
    
    marcarseleccion Me.txtReClave

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub
