VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProvCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Facturas de Compras"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   11790
   Begin VB.TextBox txtObservaciones 
      Height          =   288
      Left            =   240
      MaxLength       =   30
      TabIndex        =   15
      Top             =   2400
      Width           =   11295
   End
   Begin VB.TextBox txtAjustes 
      Height          =   285
      Left            =   9840
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtProveedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   19
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtIVA105 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtIVA27 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtIVA21 
      Height          =   285
      Left            =   9840
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPercepcionIIBB 
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtPercepcionIVA 
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtNoGravado 
      Height          =   285
      Left            =   7920
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   6000
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtNroComprob 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPrefijo 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cboComprobante 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpFechaComprobante 
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39072
   End
   Begin MSComCtl2.DTPicker dtpFechaVencimiento 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39072
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   11400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   35
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Index           =   12
      Left            =   4080
      TabIndex        =   34
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ajustes"
      Height          =   195
      Index           =   9
      Left            =   9840
      TabIndex        =   33
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe Neto"
      Height          =   195
      Index           =   14
      Left            =   7920
      TabIndex        =   32
      Top             =   840
      Width           =   915
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA 10.5%"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   31
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA 27%"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   30
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "IVA 21%"
      Height          =   195
      Index           =   5
      Left            =   9840
      TabIndex        =   29
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Perc IIBB"
      Height          =   195
      Index           =   4
      Left            =   6000
      TabIndex        =   28
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Perc IVA"
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   27
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "No Gravado"
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   26
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comprobante"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   24
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Prefijo"
      Height          =   195
      Index           =   11
      Left            =   2160
      TabIndex        =   23
      Top             =   840
      Width           =   435
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total"
      Height          =   195
      Index           =   13
      Left            =   6000
      TabIndex        =   22
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   9840
      TabIndex        =   21
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   870
   End
End
Attribute VB_Name = "frmProvCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private empresa As New clsDAOEmpresa
Private proveedor As New clsDAOProveedor
Private comprobante As New clsDAOComprobante
Private parametro As New clsDAOParametros

Private tipoComprobante As String

Private comprobantes As Collection

Private Sub cleanform()

    Me.cboComprobante.Clear
    
    Set proveedor = New clsDAOProveedor
    If proveedorglobal.proveedorID > 0 Then Set proveedor = proveedorglobal
    Me.txtProveedor.Text = proveedor.textSearch
    
    fillComprobantes
    
    Me.dtpFechaComprobante.Value = Date
    Me.dtpFechaVencimiento.Value = Date
    
    Me.txtPrefijo.Text = ""
    Me.txtNroComprob.Text = ""
    Me.txtTotal.Text = ""
    Me.txtNeto.Text = ""
    Me.txtIVA21.Text = ""
    Me.txtIVA105.Text = ""
    Me.txtIVA27.Text = ""
    Me.txtPercepcionIVA.Text = ""
    Me.txtPercepcionIIBB.Text = ""
    Me.txtNoGravado.Text = ""
    Me.txtAjustes.Text = ""
    
    Me.cmdIngresar.Enabled = True
    
    Me.cmdImprimir.Tag = 0

End Sub

Private Sub findComprobante()
Dim proveedormov As New clsDAOProveedorMov

    If Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex) < 0 Then Exit Sub

    proveedormov.proveedorID = proveedor.proveedorID
    proveedormov.comprobanteID = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
    proveedormov.prefijo = Val(Me.txtPrefijo.Text)
    proveedormov.nroComprobante = Val(Me.txtNroComprob.Text)
    proveedormov.findByComprobante db
    
    Me.cmdImprimir.Tag = proveedormov.proveedormovimientoID
    
    Me.cmdIngresar.Enabled = IIf(proveedormov.proveedormovimientoID = 0, True, False)
    
    Me.txtPrefijo.Text = proveedormov.prefijo
    Me.txtNroComprob.Text = proveedormov.nroComprobante
    Me.txtTotal.Text = Format(Abs(proveedormov.total), "0.00")
    Me.txtNeto.Text = Format(Abs(proveedormov.neto), "0.00")
    Me.txtIVA21.Text = Format(Abs(proveedormov.importeIva1), "0.00")
    Me.txtIVA105.Text = Format(Abs(proveedormov.importeIva2), "0.00")
    Me.txtIVA27.Text = Format(Abs(proveedormov.importeIva3), "0.00")
    Me.txtPercepcionIIBB.Text = Format(Abs(proveedormov.percepcionIIBB), "0.00")
    Me.txtPercepcionIVA.Text = Format(Abs(proveedormov.percepcionIva), "0.00")
    Me.txtNoGravado.Text = Format(Abs(proveedormov.gastosNoGravados), "0.00")
    Me.txtAjustes.Text = Format(Abs(proveedormov.ajustes), "0.00")
    
End Sub

Private Sub fillComprobantes()

    tipoComprobante = "C"

    If proveedor.posicionIva = 1 Or proveedor.posicionIva = 4 Then tipoComprobante = "A"
    
    comprobante.fillComboTipo Me.cboComprobante, cntTTCompras, db, , tipoComprobante
    
End Sub

Private Sub cmdImprimir_Click()
Dim ctlImp As New clsCtlImpresion

    If Me.cmdImprimir.Tag = 0 Then Exit Sub
    
    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpConsulta, "rptCompraProv", db.sconection, , Array(Array("proveedormovimientoID", Me.cmdImprimir.Tag))
    
    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdIngresar_Click()
Dim proveedormov As New clsDAOProveedorMov

Dim factor As Integer

Dim ctlCom As New clsCtlCompras

    If proveedor.proveedorID = 0 Then Exit Sub
    If Val(Me.txtTotal.Text) = 0 Then Exit Sub

    If Abs(Abs(Val(Me.txtTotal.Text)) - Abs(Val(Me.txtNeto.Text) + Val(Me.txtIVA21.Text) + Val(Me.txtIVA105.Text) + Val(Me.txtIVA27.Text) + Val(Me.txtPercepcionIVA.Text) + Val(Me.txtPercepcionIIBB.Text) + Val(Me.txtNoGravado.Text) + Val(Me.txtAjustes.Text))) > 0.15 Then
        MsgBox "ERROR: Totales INCONSISTENTES"
        Exit Sub
    End If
    
    factor = IIf(comprobante.debita = 0, 1, -1)
    
    proveedormov.proveedorID = proveedor.proveedorID
    proveedormov.comprobanteID = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
    proveedormov.prefijo = Val(Me.txtPrefijo.Text)
    proveedormov.nroComprobante = Val(Me.txtNroComprob.Text)
    proveedormov.empresaID = empresa.id
    proveedormov.fechaComprobante = Me.dtpFechaComprobante.Value
    proveedormov.fechaVencimiento = Me.dtpFechaVencimiento.Value
    proveedormov.total = Val(Me.txtTotal.Text) * factor
    proveedormov.neto = Val(Me.txtNeto.Text) * factor
    proveedormov.importeIva1 = Val(Me.txtIVA21.Text) * factor
    proveedormov.importeIva2 = Val(Me.txtIVA105.Text) * factor
    proveedormov.importeIva3 = Val(Me.txtIVA27.Text) * factor
    proveedormov.percepcionIva = Val(Me.txtPercepcionIVA.Text) * factor
    proveedormov.percepcionIIBB = Val(Me.txtPercepcionIIBB.Text) * factor
    proveedormov.gastosNoGravados = Val(Me.txtNoGravado.Text) * factor
    proveedormov.ajustes = Val(Me.txtAjustes.Text) * factor
    If tipoComprobante = "C" Then proveedormov.monotributo = 1
    proveedormov.observaciones = Me.txtObservaciones.Text
    
    If Not ctlCom.saveCompra(proveedormov, db) Then
        MsgBox "ERROR: No pudo REGISTRARSE"
        Exit Sub
    End If
    
    Me.cmdImprimir.Tag = proveedormov.proveedormovimientoID
    
    cmdImprimir_Click
    
    cleanform
    
End Sub

Private Sub cmdLimpiar_Click()

    cleanform
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()

    empresa.findLast db
    parametro.findLast db
    
    Set comprobantes = comprobante.collectionAll(db)
    
End Sub

Private Sub Form_Load()

    cleanform

End Sub

Private Sub txtAjustes_GotFocus()

    marcarseleccion Me.txtAjustes
    
End Sub

Private Sub txtIVA105_GotFocus()

    marcarseleccion Me.txtIVA105
    
End Sub

Private Sub txtIVA105_LostFocus()

    Me.txtIVA105.Text = Format(Val(Me.txtIVA105.Text), "0.00")
    
End Sub

Private Sub txtIVA21_GotFocus()

    marcarseleccion Me.txtIVA21
    
End Sub

Private Sub txtIVA21_LostFocus()

    Me.txtIVA21.Text = Format(Me.txtIVA21.Text, "0.00")
    
End Sub

Private Sub txtIVA27_GotFocus()

    marcarseleccion Me.txtIVA27

End Sub

Private Sub txtIVA27_LostFocus()

    Me.txtIVA27.Text = Format(Val(Me.txtIVA27.Text), "0.00")

End Sub

Private Sub txtNeto_GotFocus()

    marcarseleccion Me.txtNeto
    
End Sub

Private Sub txtNeto_LostFocus()

    Me.txtNeto.Text = Format(Val(Me.txtNeto.Text), "0.00")
    
End Sub

Private Sub txtNoGravado_GotFocus()

    marcarseleccion Me.txtNoGravado
    
End Sub

Private Sub txtNoGravado_LostFocus()

    Me.txtNoGravado.Text = Format(Val(Me.txtNoGravado.Text), "0.00")

End Sub

Private Sub txtNroComprob_GotFocus()

    marcarseleccion Me.txtNroComprob
    
End Sub

Private Sub txtNroComprob_LostFocus()

    findComprobante
    
End Sub

Private Sub txtObservaciones_GotFocus()

    marcarseleccion Me.txtObservaciones
    
End Sub

Private Sub txtPercepcionIIBB_GotFocus()

    marcarseleccion Me.txtPercepcionIIBB
    
End Sub

Private Sub txtPercepcionIIBB_LostFocus()

    Me.txtPercepcionIIBB.Text = Format(Val(Me.txtPercepcionIIBB.Text), "0.00")

End Sub

Private Sub txtPercepcionIVA_GotFocus()

    marcarseleccion Me.txtPercepcionIVA
    
End Sub

Private Sub txtPercepcionIVA_LostFocus()

    Me.txtPercepcionIVA.Text = Format(Val(Me.txtPercepcionIVA.Text), "0.00")

End Sub

Private Sub txtPrefijo_GotFocus()

    marcarseleccion Me.txtPrefijo
    
End Sub

Private Sub txtPrefijo_LostFocus()

    findComprobante
    
End Sub

Private Sub txtProveedor_GotFocus()

    marcarseleccion Me.txtProveedor
    
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
Dim ctlSrc As New clsCtlSearch

    ctlSrc.formSearch proveedor, KeyAscii, "Proveedores", db
    
    Me.txtProveedor.Text = proveedor.textSearch
    KeyAscii = 0
    
    Set proveedorglobal = proveedor
    
    fillComprobantes
    
End Sub

Private Sub txtTotal_GotFocus()

    marcarseleccion Me.txtTotal
    
End Sub

Private Sub txtTotal_LostFocus()

    Me.txtTotal.Text = Format(Me.txtTotal.Text, "0.00")
    
End Sub

