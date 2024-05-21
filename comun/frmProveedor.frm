VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   7845
   ClientLeft      =   840
   ClientTop       =   1890
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11790
   Begin VB.TextBox txtObservaciones 
      Height          =   1485
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   5760
      Width           =   5535
   End
   Begin VB.TextBox txtProvincia 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtLocalidad 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtContacto 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   5535
   End
   Begin VB.TextBox txtFantasia 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox txtIngresosBrutos 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox lstProveedores 
      Height          =   7275
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   360
      Width           =   5535
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtCelular 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txteMail 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   5535
   End
   Begin VB.ComboBox cboIVA 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   7320
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mebCUIT 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   34
      Top             =   5520
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Provincia"
      Height          =   195
      Index           =   14
      Left            =   4080
      TabIndex        =   33
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Localidad"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   32
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Contacto"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   31
      Top             =   4920
      Width           =   645
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Nombre Fantasía"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   30
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ingresos Brutos"
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedores"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   28
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   26
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Celular"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      Height          =   195
      Index           =   8
      Left            =   2160
      TabIndex        =   21
      Top             =   3720
      Width           =   435
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "C.U.I.T."
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Posición I.V.A."
      Height          =   195
      Index           =   10
      Left            =   4080
      TabIndex        =   18
      Top             =   4320
      Width           =   1035
   End
End
Attribute VB_Name = "frmProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private proveedor As New clsDAOProveedor

Public Sub fillForm()

On Error Resume Next
    
    Me.mebCUIT.Text = "00-00000000-0"
    
    With proveedor
        Me.txtCodigo.Text = .proveedorID
        Me.txtNombre.Text = .razonSocial
        Me.txtFantasia.Text = .nombreFantasia
        Me.txtDomicilio.Text = .domicilio
        Me.txtLocalidad.Text = .localidad
        Me.txtProvincia.Text = .provincia
        Me.mebCUIT.Text = .cuit
        Me.txtFax.Text = .fax
        Me.txtTelefono.Text = .telefono
        Me.txtCelular.Text = .celular
        Me.txteMail.Text = .email
        Me.txtContacto.Text = .contacto
        Me.txtObservaciones.Text = .observaciones
        If .posicionIva = 0 Then
            Me.cboIVA.Text = "Consum.Final"
        Else
            If .posicionIva = 1 Then Me.cboIVA.Text = "Resp.Inscripto"
            If .posicionIva = 2 Then Me.cboIVA.Text = "Consum.Final"
            If .posicionIva = 3 Then Me.cboIVA.Text = "Monotributista"
            If .posicionIva = 4 Then Me.cboIVA.Text = "Resp.No Inscripto"
            If .posicionIva = 5 Then Me.cboIVA.Text = "Exportación"
            If .posicionIva = 6 Then Me.cboIVA.Text = "Exento"
        End If
    End With
    
End Sub

Private Sub cboIVA_Click()
    
    proveedor.posicionIva = Me.cboIVA.ItemData(Me.cboIVA.ListIndex)
    
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
    
    With proveedor
        .save db
        .fillList Me.lstProveedores, db
    End With
    
    cmdLimpiar_Click
    
End Sub

Private Sub cmdLimpiar_Click()

    proveedor.newID True, db
    
    fillForm

End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set proveedor = Nothing
    
End Sub

Private Sub lstProveedores_Click()

    If Me.lstProveedores.ListIndex < 0 Then Exit Sub
    
    proveedor.proveedorID = Me.lstProveedores.ItemData(Me.lstProveedores.ListIndex)
    proveedor.findByPrimaryKey db
    
    fillForm

End Sub

Private Sub mebCUIT_GotFocus()

    Me.mebCUIT.SelStart = 0
    Me.mebCUIT.SelLength = Len(Me.mebCUIT.Text)
    
End Sub

Private Sub mebCUIT_LostFocus()
Dim strCliente As String
    
    With proveedor
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

    proveedor.celular = Me.txtCelular.Text
    
End Sub

Private Sub txtCodigo_GotFocus()

    marcarseleccion Me.txtCodigo
    
End Sub

Private Sub txtCodigo_LostFocus()

    proveedor.proveedorID = Me.txtCodigo.Text
    
End Sub

Private Sub txtContacto_GotFocus()

    marcarseleccion Me.txtContacto
    
End Sub

Private Sub txtContacto_LostFocus()

    proveedor.contacto = Me.txtContacto.Text
    
End Sub

Private Sub txtDomicilio_GotFocus()

    marcarseleccion Me.txtDomicilio
    
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)

    If InStr(cntProhibidos, Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtDomicilio_LostFocus()

    proveedor.domicilio = Me.txtDomicilio.Text
    
End Sub

Private Sub txteMail_GotFocus()

    marcarseleccion Me.txteMail
    
End Sub

Private Sub txteMail_LostFocus()

    proveedor.email = Me.txteMail.Text
    
End Sub

Private Sub txtFantasia_GotFocus()

    marcarseleccion Me.txtFantasia
    
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)

    If InStr(cntProhibidos, Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtFantasia_LostFocus()

    proveedor.nombreFantasia = Me.txtFantasia.Text
    
End Sub

Private Sub txtFax_GotFocus()

    marcarseleccion Me.txtFax
    
End Sub

Private Sub txtFax_LostFocus()

    proveedor.fax = Me.txtFax.Text
    
End Sub

Private Sub txtLocalidad_GotFocus()

    marcarseleccion Me.txtLocalidad
    
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)

    If InStr(cntProhibidos, Chr(KeyAscii)) Then KeyAscii = 0
    
End Sub

Private Sub txtLocalidad_LostFocus()

    proveedor.localidad = Me.txtLocalidad.Text
    
End Sub

Private Sub txtNombre_GotFocus()

    marcarseleccion Me.txtNombre
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)

    If InStr(cntProhibidos, Chr(KeyAscii)) Then KeyAscii = 0

End Sub

Private Sub txtNombre_LostFocus()

    proveedor.razonSocial = Me.txtNombre.Text
    
End Sub

Private Sub txtObservaciones_GotFocus()

    marcarseleccion Me.txtObservaciones
    
End Sub

Private Sub txtObservaciones_LostFocus()

    proveedor.observaciones = Me.txtObservaciones.Text
    
End Sub

Private Sub txtProvincia_GotFocus()

    marcarseleccion Me.txtProvincia
    
End Sub

Private Sub txtProvincia_KeyPress(KeyAscii As Integer)

    If InStr(cntProhibidos, Chr(KeyAscii)) Then KeyAscii = 0
    
End Sub

Private Sub txtProvincia_LostFocus()

    proveedor.provincia = Me.txtProvincia.Text
    
End Sub

Private Sub txtTelefono_GotFocus()

    marcarseleccion Me.txtTelefono
    
End Sub

Private Sub txtTelefono_LostFocus()

    proveedor.telefono = Me.txtTelefono.Text
    
End Sub

Private Sub Form_Load()

    fillComboIva Me.cboIVA
    
    With proveedor
        .fillList Me.lstProveedores, db
        .newID True, db
    End With
    
    fillForm
    
End Sub

