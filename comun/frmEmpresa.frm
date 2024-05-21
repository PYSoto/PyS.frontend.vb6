VERSION 5.00
Begin VB.Form frmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Empresa"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7935
   Begin VB.TextBox txtUbicacion 
      Height          =   285
      Left            =   6000
      TabIndex        =   18
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtRSocial 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   7455
   End
   Begin VB.TextBox txtCondicionIVA 
      Height          =   285
      Left            =   6000
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtInicioActividades 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtIngBrutos 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtCUIT 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
      Height          =   195
      Index           =   6
      Left            =   6000
      TabIndex        =   19
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Condición IVA"
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   16
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Inicio Actividades"
      Height          =   195
      Index           =   8
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ing. Brutos"
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "CUIT"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   12
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Nombre Empresa"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objEmpresa As New clsDAOEmpresa

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()

    objEmpresa.findLast db
     
    llenarFormulario
    
End Sub

Public Sub llenarFormulario()
    
    With objEmpresa
        Me.txtNombre.Text = .nombre
        Me.txtRSocial.Text = .rsocial
        Me.txtDomicilio.Text = .domicilio
        Me.txtTelefono.Text = .telf
        Me.txtCUIT.Text = .cuit
        Me.txtIngBrutos.Text = .ingbrutos
        Me.txtInicioActividades.Text = .inicioactividades
        Me.txtCondicionIVA.Text = .condicioniva
        Me.txtUbicacion.Text = .ubicacion
    End With

End Sub

Private Sub cmdGuardar_Click()

    With objEmpresa
        .nombre = Me.txtNombre.Text
        .rsocial = Me.txtRSocial.Text
        .domicilio = Me.txtDomicilio.Text
        .telf = Me.txtTelefono.Text
        .cuit = Me.txtCUIT.Text
        .ingbrutos = Me.txtIngBrutos.Text
        .inicioactividades = Me.txtInicioActividades.Text
        .condicioniva = Me.txtCondicionIVA.Text
        .ubicacion = Me.txtUbicacion.Text
        
        .save db
    End With

End Sub

Private Sub txtCondicionIVA_GotFocus()

    marcarseleccion Me.txtCondicionIVA
    
End Sub

Private Sub txtCUIT_GotFocus()

    marcarseleccion Me.txtCUIT
    
End Sub

Private Sub txtDomicilio_GotFocus()

    marcarseleccion Me.txtDomicilio
    
End Sub

Private Sub txtIngBrutos_GotFocus()

    marcarseleccion Me.txtIngBrutos
    
End Sub

Private Sub txtInicioActividades_GotFocus()

    marcarseleccion Me.txtInicioActividades
    
End Sub

Private Sub txtNombre_GotFocus()

    marcarseleccion Me.txtNombre
    
End Sub

Private Sub txtRSocial_GotFocus()

    marcarseleccion Me.txtRSocial
    
End Sub

Private Sub txtTelefono_GotFocus()

    marcarseleccion Me.txtTelefono
    
End Sub

Private Sub txtUbicacion_GotFocus()

    marcarseleccion Me.txtUbicacion
    
End Sub
