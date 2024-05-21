VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   5760
   ClientLeft      =   840
   ClientTop       =   1890
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9855
   Begin VB.CheckBox chkFacturable 
      Caption         =   "Facturable"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboDescuento 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtProvincia 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtLocalidad 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtNroDocumento 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtTipoDocumento 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtLimiteCredito 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ListBox lstClientes 
      Height          =   5325
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtCelular 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txteMail 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.ComboBox cboIVA 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   5280
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mebCUIT 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Provincia"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Localidad"
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Número Documento"
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Documento"
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Límite Crédito"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Clientes"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   27
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Celular"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   3840
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "C.U.I.T."
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Posición I.V.A."
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1035
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objCliente As New clsDAOClientes

Private Sub cboDescuento_Click()

    objCliente.descuento = Val(Me.cboDescuento.Text)
    
End Sub

Private Sub cboIVA_Click()
    
    objCliente.posicion = Me.cboIVA.ItemData(Me.cboIVA.ListIndex)
    
End Sub

Private Sub chkFacturable_Click()

    objCliente.facturable = Me.chkFacturable.Value
    
End Sub

Private Sub cmdGrabar_Click()
    
    If txtNombre.Text = "" Then
        MsgBox "Debe ingresar el nombre"
        txtNombre.SetFocus
        Exit Sub
    End If
    If mebCUIT.Text = "__-________-_" Then
        MsgBox "ERROR: Debe ingresar Nro.CUIT"
        mebCUIT.SetFocus
        Exit Sub
    End If
    
    With objCliente
        .save db
        .fillList Me.lstClientes, db
    End With
    
    Call cmdCancelar_Click
    
End Sub

Private Sub cmdCancelar_Click()

    objCliente.newID True, db
    
    llenarFormulario
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objCliente = Nothing
    
End Sub

Private Sub lstClientes_Click()

    If Me.lstClientes.ListIndex < 0 Then Exit Sub
    
    objCliente.findByCodigo Me.lstClientes.ItemData(Me.lstClientes.ListIndex), db
    
    llenarFormulario
    
End Sub

Private Sub mebCUIT_GotFocus()

    Me.mebCUIT.SelStart = 0
    Me.mebCUIT.SelLength = Len(Me.mebCUIT.Text)
    
End Sub

Private Sub mebCUIT_LostFocus()
Dim strCliente As String
    
    With objCliente
        .cuit = Me.mebCUIT.Text
        
        If .existByCUIT(db, strCliente) Then
            MsgBox "El CUIT ya existe (" & strCliente & ")", , "ERROR"
            Me.mebCUIT.SetFocus
        End If
        
        If Me.mebCUIT.Text <> "00-00000000-0" Then If Not modCUIT.validaCUIT(Replace(Me.mebCUIT.Text, "-", "")) Then MsgBox "ERROR: CUIT NO Válido"
    End With
    
End Sub

Private Sub txtCelular_GotFocus()

    marcarseleccion Me.txtCelular
    
End Sub

Private Sub txtCelular_LostFocus()

    objCliente.celular = Me.txtCelular.Text
    
End Sub

Private Sub txtCodigo_GotFocus()

    marcarseleccion Me.txtCodigo
    
End Sub

Private Sub txtCodigo_LostFocus()

    objCliente.codigo = Me.txtCodigo.Text
    
End Sub

Private Sub txtDomicilio_GotFocus()

    marcarseleccion Me.txtDomicilio
    
End Sub

Private Sub txtDomicilio_LostFocus()

    objCliente.domicilio = Me.txtDomicilio.Text
    
End Sub

Private Sub txteMail_GotFocus()

    marcarseleccion Me.txteMail
    
End Sub

Private Sub txteMail_LostFocus()

    objCliente.email = Me.txteMail.Text
    
End Sub

Private Sub txtFax_GotFocus()

    marcarseleccion Me.txtFax
    
End Sub

Private Sub txtFax_LostFocus()

    objCliente.fax = Me.txtFax.Text
    
End Sub

Private Sub txtLimiteCredito_GotFocus()

    marcarseleccion Me.txtLimiteCredito
    
End Sub

Private Sub txtLimiteCredito_LostFocus()

    objCliente.limitecredito = Val(Me.txtLimiteCredito.Text)
    
End Sub

Private Sub txtLocalidad_GotFocus()

    marcarseleccion Me.txtLocalidad
    
End Sub

Private Sub txtLocalidad_LostFocus()

    objCliente.localidad = Me.txtLocalidad.Text
    
End Sub

Private Sub txtNombre_GotFocus()

    marcarseleccion Me.txtNombre
    
End Sub

Private Sub txtNombre_LostFocus()

    objCliente.razon = Me.txtNombre.Text
    
End Sub

Private Sub txtNroDocumento_GotFocus()

    marcarseleccion Me.txtNroDocumento
    
End Sub

Private Sub txtNroDocumento_LostFocus()

    objCliente.nrodoc = Val(Me.txtNroDocumento.Text)
    
End Sub

Private Sub txtProvincia_GotFocus()

    marcarseleccion Me.txtProvincia
    
End Sub

Private Sub txtProvincia_LostFocus()

    objCliente.provincia = Me.txtProvincia.Text
    
End Sub

Private Sub txtTelefono_GotFocus()

    marcarseleccion Me.txtTelefono
    
End Sub

Private Sub txtTelefono_LostFocus()

    objCliente.tel = Me.txtTelefono.Text
    
End Sub

Private Sub Form_Load()
Dim varDescuentos As Variant

Dim intCiclo As Integer

    llenaComboIVA Me.cboIVA
    
    With objCliente
        .fillList Me.lstClientes, db
        .newID True, db
    End With
    
    llenarFormulario
    
    varDescuentos = Array(-20, -15, -10, -5, 0, 5, 10, 15, 20)
    
    Me.cboDescuento.Clear
    
    For intCiclo = LBound(varDescuentos) To UBound(varDescuentos)
        Me.cboDescuento.AddItem varDescuentos(intCiclo)
    Next intCiclo
    
End Sub

Public Sub llenarFormulario()

On Error Resume Next
    
    Me.mebCUIT.Text = "00-00000000-0"
    
    With objCliente
        Me.txtCodigo.Text = .codigo
        Me.txtNombre.Text = .razon
        Me.txtDomicilio.Text = .domicilio
        Me.txtLocalidad.Text = .localidad
        Me.txtProvincia.Text = .provincia
        Me.mebCUIT.Text = .cuit
        Me.txtFax.Text = .fax
        Me.txtTelefono.Text = .tel
        Me.txtCelular.Text = .celular
        Me.txteMail.Text = .email
        Me.txtLimiteCredito.Text = Format(.limitecredito, "0.00")
        Me.txtTipoDocumento.Text = .tipodoc
        Me.txtNroDocumento.Text = .nrodoc
        Me.cboDescuento.Text = .descuento
        Me.chkFacturable.Value = .facturable
        If .posicion = 0 Then
            Me.cboIVA.Text = "Consum.Final"
        Else
            If .posicion = 1 Then Me.cboIVA.Text = "Resp.Inscripto"
            If .posicion = 2 Then Me.cboIVA.Text = "Consum.Final"
            If .posicion = 3 Then Me.cboIVA.Text = "Monotributista"
            If .posicion = 4 Then Me.cboIVA.Text = "Resp.No Inscripto"
            If .posicion = 5 Then Me.cboIVA.Text = "Exportación"
            If .posicion = 6 Then Me.cboIVA.Text = "Exento"
        End If
    End With
    
End Sub

Private Sub txtTipoDocumento_GotFocus()

    marcarseleccion Me.txtTipoDocumento
    
End Sub

Private Sub txtTipoDocumento_LostFocus()

    objCliente.tipodoc = Me.txtTipoDocumento.Text
    
End Sub
