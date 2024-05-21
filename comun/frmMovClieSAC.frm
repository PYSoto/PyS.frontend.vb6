VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovClieSAC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas de Otro Negocio"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   9855
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox cboNegocio 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtIVA21 
      Height          =   285
      Left            =   6000
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtNroComprob 
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPrefijo 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cboTipoComprobante 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77660161
      CurrentDate     =   39072
   End
   Begin MSComCtl2.DTPicker dtpVencimiento 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77660161
      CurrentDate     =   39072
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Negocio"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Neto"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   20
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA"
      Height          =   195
      Index           =   5
      Left            =   6000
      TabIndex        =   19
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   18
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comprobante"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   6000
      TabIndex        =   16
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total"
      Height          =   195
      Left            =   7920
      TabIndex        =   15
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   870
   End
End
Attribute VB_Name = "frmMovClieSAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objCliente As New clsDAOClientes

Private objTiposComprob As New clsDAOTiposComprob

Private objMovclie As New clsDAOMovclie

Private objNegocio As New clsDAONegocio

Private vTipoCompro As String

Private Sub cboNegocio_Click()

    If Me.cboNegocio.ListIndex < 0 Then Exit Sub
    
    configuraCliente
    
End Sub

Private Sub cboTipoComprobante_Click()

    If Me.cboTipoComprobante.ListIndex < 0 Then Exit Sub
    
    objTiposComprob.codigo = Me.cboTipoComprobante.ItemData(Me.cboTipoComprobante.ListIndex)
    objTiposComprob.findByPrimaryKey db
    
    Me.txtPrefijo.Text = objTiposComprob.puntovta
    
    fillForm
    
End Sub

Private Sub cmdCancelar_Click()

    Set objMovclie = New clsDAOMovclie
    
    Me.dtpFecha.Value = Date
    Me.dtpVencimiento.Value = Date
    Me.txtPrefijo.Text = ""
    Me.txtNroComprob.Text = ""
    Me.txtTotal.Text = ""
    Me.txtNeto.Text = ""
    Me.txtIVA21.Text = ""
    
End Sub

Private Sub cmdEliminar_Click()
Dim lngResp As Long

    lngResp = MsgBox("Está SEGURO ?", vbYesNo, "Eliminar COMPROBANTE")
    
    If lngResp = vbNo Then Exit Sub

    With objMovclie
        If .clave = 0 Then Exit Sub
        
        .delete db
    End With
    
    fillForm
    
End Sub

Private Sub cmdIngresar_Click()

    If Me.cboNegocio.ListIndex < 0 Then Exit Sub
    If Me.cboTipoComprobante.ListIndex < 0 Then Exit Sub
    
    If Val(Me.txtTotal.Text) = 0 Then
        MsgBox "ERROR: SIN Importe"
        Exit Sub
    End If
    
    Set objMovclie = New clsDAOMovclie
    
    With objMovclie
        .findByComprobante Me.cboNegocio.ItemData(Me.cboNegocio.ListIndex), Me.cboTipoComprobante.ItemData(Me.cboTipoComprobante.ListIndex), Val(Me.txtPrefijo.Text), Val(Me.txtNroComprob.Text), db
        
        If .clave <> 0 Then
            MsgBox "ERROR: Comprobante CARGADO"
            Exit Sub
        End If
        
        .negid = Me.cboNegocio.ItemData(Me.cboNegocio.ListIndex)
        .cgoclie = objCliente.codigo
        .cgocomprob = Me.cboTipoComprobante.ItemData(Me.cboTipoComprobante.ListIndex)
        .fechacomprob = Me.dtpFecha.Value
        .fechavenc = Me.dtpVencimiento.Value
        .prefijo = Val(Me.txtPrefijo.Text)
        .nroComprob = Val(Me.txtNroComprob.Text)
        .importe = Val(Me.txtTotal.Text)
        .neto = Val(Me.txtNeto.Text)
        .montoiva = Val(Me.txtIVA21.Text)
        .tipocompro = vTipoCompro
    
        .add db
    End With

    Call cmdCancelar_Click
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpFecha_Change()

    Me.dtpVencimiento.Value = Me.dtpFecha.Value
    
End Sub

Private Sub Form_Activate()

    objNegocio.fillComboExcluidoLocal Me.cboNegocio
    
End Sub

Private Sub Form_Load()

    Call cmdCancelar_Click
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)

    objCliente.formBuscar frmBuscar, objCliente, KeyAscii, "Clientes"
    
    Me.txtCliente.Text = objCliente.descripcionBuscar
    KeyAscii = 0
    
    configuraCliente

End Sub

Private Sub txtIVA21_GotFocus()

    marcarseleccion Me.txtIVA21
    
End Sub

Private Sub txtNeto_GotFocus()

    marcarseleccion Me.txtNeto
    
End Sub

Private Sub txtNeto_KeyPress(KeyAscii As Integer)
Dim objPa As New clsDAOParametros

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        objPa.findLast
        Me.txtIVA21.Text = Format(Val(Me.txtNeto.Text) * (objPa.iva1 / 100), "0.00")
        Me.txtTotal.Text = Format(Val(Me.txtNeto.Text) + Val(Me.txtIVA21.Text), "0.00")
    End If
    
End Sub

Private Sub txtNeto_LostFocus()

    txtNeto_KeyPress 13
    
End Sub

Private Sub txtNroComprob_GotFocus()

    marcarseleccion Me.txtNroComprob
    
End Sub

Private Sub txtNroComprob_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        txtNroComprob_LostFocus
    End If
    
End Sub

Private Sub txtNroComprob_LostFocus()

    fillForm
    
End Sub

Private Sub txtPrefijo_GotFocus()

    marcarseleccion Me.txtPrefijo
    
End Sub

Private Sub txtTotal_GotFocus()

    marcarseleccion Me.txtTotal
    
End Sub

Private Sub configuraCliente()

    If Me.cboNegocio.ListIndex < 0 Then Exit Sub

    vTipoCompro = "B"
    
    If objCliente.posicion = 1 Or objCliente.posicion = 4 Then vTipoCompro = "A"
    
    objTiposComprob.fillComboVentas Me.cboTipoComprobante, vTipoCompro, 0, Me.cboNegocio.ItemData(Me.cboNegocio.ListIndex), db
    
End Sub

Private Sub fillForm()

    If Me.cboTipoComprobante.ListIndex < 0 Then Exit Sub
    
    With objMovclie
        .findByComprobante Me.cboNegocio.ItemData(Me.cboNegocio.ListIndex), Me.cboTipoComprobante.ItemData(Me.cboTipoComprobante.ListIndex), Val(Me.txtPrefijo.Text), Val(Me.txtNroComprob.Text), db
    
        Me.dtpFecha.Value = .fechacomprob
        Me.dtpVencimiento.Value = .fechavenc
        Me.txtNeto.Text = Format(.neto, "0.00")
        Me.txtIVA21.Text = Format(.montoiva, "0.00")
        Me.txtTotal.Text = Format(.importe, "0.00")
    End With
    
End Sub
