VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProvPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Pagos"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11805
   Begin VB.TextBox txtSaldo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   9840
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtTitular 
      Height          =   285
      Left            =   7920
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminarValor 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgregarValor 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtImporte 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox cboValor 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox txtTotalValores 
      Height          =   288
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtProveedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   18
      Top             =   6840
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdFacturas 
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   9840
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboComprobante 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39072
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   11040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpEmision 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   31260673
      CurrentDate     =   39074
   End
   Begin MSFlexGridLib.MSFlexGrid grdValores 
      Height          =   2535
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4471
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker dtpVencimiento 
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39074
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobantes"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   " Valores "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      Height          =   195
      Index           =   24
      Left            =   9840
      TabIndex        =   31
      Top             =   3600
      Width           =   465
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Titular"
      Height          =   195
      Index           =   23
      Left            =   7920
      TabIndex        =   30
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Index           =   22
      Left            =   6000
      TabIndex        =   29
      Top             =   3600
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Emisión"
      Height          =   195
      Index           =   21
      Left            =   4080
      TabIndex        =   28
      Top             =   3600
      Width           =   540
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Index           =   20
      Left            =   2160
      TabIndex        =   27
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      Height          =   195
      Index           =   18
      Left            =   4080
      TabIndex        =   26
      Top             =   3000
      Width           =   525
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total Valores"
      Height          =   195
      Index           =   19
      Left            =   9840
      TabIndex        =   24
      Top             =   3000
      Width           =   930
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11640
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comprobante"
      Height          =   195
      Left            =   6000
      TabIndex        =   21
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total"
      Height          =   195
      Left            =   9840
      TabIndex        =   20
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   7920
      TabIndex        =   19
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmProvPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private proveedor As New clsDAOProveedor
Private proveedormov As clsDAOProveedorMov
Private empresa As New clsDAOEmpresa
Private parametro As New clsDAOParametros

Private selectedValor As Boolean

Private row As Integer

Private lastValor As Long

Private comprobantes As Collection
Private valores As Collection
Private proveedormovs As Collection
Private valormovs As Collection

Private Const cntPago = 6

Private Sub cleanform()
Dim comprobante As New clsDAOComprobante

    Me.dtpFecha.Value = Date
    Me.dtpEmision.Value = Date
    Me.dtpVencimiento = Date
    Me.cboComprobante.Clear
    
    Set proveedor = New clsDAOProveedor
    If proveedorglobal.proveedorID > 0 Then
        Set proveedor = proveedorglobal
        comprobante.fillComboTipo Me.cboComprobante, cntTTPagos, db
    End If

    Me.txtProveedor.Text = proveedor.textSearch
    Me.grdFacturas.Rows = 1
    
    Me.txtTotal.Locked = False
    
    row = 0
    
    Set proveedormovs = Nothing
    Set valormovs = New Collection
    Me.grdValores.Rows = 1
    
    Me.txtTotalValores.Text = ""
    
End Sub

Private Sub fillFacturasPendientes()
Dim comprobanteloc As clsDAOComprobante
Dim proveedormovloc As New clsDAOProveedorMov

    Me.grdFacturas.Rows = 1
    
    If proveedor.proveedorID = 0 Then Exit Sub
    
    Set proveedormovs = proveedormovloc.collectionPendientesByProveedorID(proveedor.proveedorID, db)
    
    For Each proveedormovloc In proveedormovs
        Set comprobanteloc = comprobantes("k." & proveedormovloc.comprobanteID)
        Me.grdFacturas.AddItem modGrid.array2itemGrid(Array(comprobanteloc.descripcion, modConv.formatNumComprobante(proveedormovloc.prefijo, proveedormovloc.nroComprobante), proveedormovloc.fechaComprobante, Format(proveedormovloc.total, "0.00"), Format(proveedormovloc.totalCancelado, "0.00"), Format(proveedormovloc.total - proveedormovloc.totalCancelado, "0.00"), "0.00"))
        Me.grdFacturas.RowData(Me.grdFacturas.Rows - 1) = proveedormovloc.proveedormovimientoID
    Next

    If Me.grdFacturas.Rows = 1 Then MsgBox "No existen comprobantes pendientes"

End Sub

Private Function totalFacturas() As Currency
Dim proveedormovloc As clsDAOProveedorMov

Dim total As Currency

    total = 0
    For Each proveedormovloc In proveedormovs
        total = total + proveedormovloc.pago
    Next

    totalFacturas = total

End Function

Private Sub fillValores()
Dim valormov As clsDAOValorMov
Dim valor As New clsDAOValor

Dim total As Currency

    total = 0

    Me.grdValores.Rows = 1
    Me.grdValores.Redraw = False
    For Each valormov In valormovs
        Set valor = valores("k." & valormov.valorID)
        Me.grdValores.AddItem modGrid.array2itemGrid(Array(valor.comboText, Format(valormov.importe, "0.00"), valormov.nroComprobante, valormov.fechaEmision, valormov.fechaVencimiento, valormov.titular, valormov.banco))
        Me.grdValores.RowData(Me.grdValores.Rows - 1) = valormov.keyCollection
        
        total = total + valormov.importe
    Next
    Me.grdValores.Redraw = True
    
    Me.txtTotalValores.Text = Format(total, "0.00")

End Sub

Private Sub cboComprobante_Click()
Dim comprobante As clsDAOComprobante
    
    Me.grdFacturas.Rows = 1
    
    If Me.cboComprobante.ListIndex < 0 Then Exit Sub
    
    If comprobantes Is Nothing Then
        Set comprobante = New clsDAOComprobante
        comprobante.comprobanteID = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
        comprobante.findByPrimaryKey db
    Else
        Set comprobante = comprobantes("k." & Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex))
    End If
    
    Me.txtTotal.Text = "0.00"
    Me.txtTotal.Locked = False
    Me.cmdAgregarValor.Enabled = True
    Me.cmdEliminarValor.Enabled = True
    
    If comprobante.aplicaPendiente <> 0 Then
        fillFacturasPendientes
        Me.txtTotal.Locked = True
    End If
    
    If comprobante.aplicacion <> 0 Then
        Me.grdValores.Rows = 1
        Set valormovs = New Collection
        Me.cmdAgregarValor.Enabled = False
        Me.cmdEliminarValor.Enabled = False
        Me.txtTotalValores.Text = ""
        Me.txtTotal.Locked = True
    End If
    
End Sub

Private Sub cmdAgregarValor_Click()
Dim valormov As New clsDAOValorMov

    If Me.cboValor.ListIndex < 0 Then Exit Sub
    If Val(Me.txtImporte.Text) = 0 Then Exit Sub
    
    If selectedValor Then Set valormov = valormovs("k." & Me.grdValores.RowData(Me.grdValores.row))
    valormov.valorID = Me.cboValor.ItemData(Me.cboValor.ListIndex)
    valormov.proveedorID = proveedor.proveedorID
    valormov.importe = Val(Me.txtImporte.Text)
    valormov.fechaRegistracion = Me.dtpFecha.Value
    valormov.comprobanteID = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
    valormov.nroComprobante = Val(Me.txtNumero.Text)
    valormov.titular = Me.txtTitular.Text
    valormov.banco = Me.txtBanco.Text
    valormov.fechaEmision = Me.dtpEmision.Value
    valormov.fechaVencimiento = Me.dtpVencimiento.Value
    valormov.negocioID = gEmpresa.negid
    If Not selectedValor Then
        lastValor = lastValor + 1
        valormov.keyCollection = lastValor
    
        valormovs.add valormov, "k." & lastValor
    End If
    
    fillValores
    
    selectedValor = False
    Me.cmdAgregarValor.Caption = "Agregar"
    
End Sub

Private Sub cmdEliminarValor_Click()

    If Me.grdValores.row < 1 Then Exit Sub
    
    valormovs.Remove "k." & Me.grdValores.RowData(Me.grdValores.row)
    
    selectedValor = False
    Me.cmdAgregarValor.Caption = "Agregar"
    
    fillValores
    
End Sub

Private Sub cmdIngresar_Click()
Dim comprobante As clsDAOComprobante
Dim proveedormovloc As clsDAOProveedorMov

Dim ctlCmp As New clsCtlCompras
Dim ctlImp As New clsCtlImpresion

    If Me.cboComprobante.ListIndex < 0 Then Exit Sub
    
    If Val(Me.txtTotal.Text) <> Val(Me.txtTotalValores.Text) Then
        MsgBox "ERROR: Diferencia entre VALORES y TOTAL"
        Exit Sub
    End If
    
    If Not proveedormovs Is Nothing Then
        For Each proveedormovloc In proveedormovs
            If Abs(proveedormovloc.totalCancelado + proveedormovloc.pago) > Abs(proveedormovloc.total) Then
                MsgBox "ERROR: Comprobante SIN Saldo SUFICIENTE"
                Exit Sub
            End If
        Next
    End If
    
    Set comprobante = comprobantes("k." & Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex))
    
    Set proveedormov = New clsDAOProveedorMov
    proveedormov.proveedorID = proveedor.proveedorID
    proveedormov.comprobanteID = comprobante.comprobanteID
    proveedormov.fechaComprobante = Me.dtpFecha.Value
    proveedormov.prefijo = comprobante.comprobanteID
    proveedormov.negocioID = gEmpresa.negid
    proveedormov.empresaID = gEmpresa.id
    If comprobante.ordenPago <> 0 Then proveedormov.nroComprobante = ctlCmp.nextOP(db)
    If comprobante.aplicaPendiente <> 0 Then If proveedormov.nroComprobante = 0 Then proveedormov.nroComprobante = ctlCmp.nextComprobante(comprobante.comprobanteID, db)
    proveedormov.total = -Val(Me.txtTotal.Text)
    proveedormov.gastosNoGravados = proveedormov.total
    
    If Not ctlCmp.savePago(proveedormov, proveedor, proveedormovs, valormovs, db) Then
        MsgBox "ERROR: No pudo REGISTRARSE"
        Exit Sub
    End If
    
    MsgBox "Comprobante GENERADO: " & modConv.formatNumComprobante(proveedormov.prefijo, proveedormov.nroComprobante)
    
    ctlImp.printReport Me.crpConsulta, "rptPagoProv", db.sconection, Array("FactPagadas", "Valores"), Array(Array("proveedormovimientoID", proveedormov.proveedormovimientoID))
    
    cleanform

End Sub

Private Sub cmdLimpiar_Click()

    cleanform
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
Dim valor As New clsDAOValor
Dim comprobante As New clsDAOComprobante
    
    valor.fillCombo Me.cboValor, db
    
    Set valores = valor.collectionAll(db)
    Set comprobantes = comprobante.collectionAll(db)
    
End Sub

Private Sub Form_Load()

    modGrid.makeGrid2 Me.grdValores, Array(Array("Concepto", 4400), Array("Importe", 1500), Array("Numero", 1000), Array("Emision", 1000), Array("Vencimiento", 1000), Array("Titular", 1000), Array("Banco", 1000)), 0, 1, flexSelectionByRow
    modGrid.makeGrid2 Me.grdFacturas, Array(Array("Comprobante", 4400), Array("Numero", 1500), Array("Fecha", 1000), Array("Total", 1000), Array("A Cta", 1000), Array("Saldo", 1000), Array("Pago", 1000)), 0, 1, flexSelectionFree
    
    empresa.findLast db
    parametro.findLast db

    cleanform
    
End Sub

Private Sub grdFacturas_Click()

    If Me.grdFacturas.Rows = 1 Then Exit Sub

    If Me.grdFacturas.Col <> 6 Then Exit Sub

    row = Me.grdFacturas.row
    
    Set proveedormov = proveedormovs("k." & Me.grdFacturas.RowData(row))

    Me.txtSaldo.Left = modGrid.leftGrid(Me.grdFacturas)
    Me.txtSaldo.Top = modGrid.topGrid(Me.grdFacturas)
    Me.txtSaldo.Width = Me.grdFacturas.ColWidth(Me.grdFacturas.Col)
    Me.txtSaldo.Height = Me.grdFacturas.CellHeight

    Me.txtSaldo.Text = Me.grdFacturas.TextMatrix(row, cntPago)

    Me.txtSaldo.Visible = True

    Me.txtSaldo.SetFocus
    
End Sub

Private Sub grdValores_Click()
Dim valor As New clsDAOValor
Dim valormov As clsDAOValorMov

    If Me.grdValores.row < 1 Then Exit Sub
    
    selectedValor = True
    Set valormov = valormovs("k." & Me.grdValores.RowData(Me.grdValores.row))
    Set valor = valores("k." & valormov.valorID)
    
    Me.txtImporte.Text = Format(valormov.importe, "0.00")
    Me.cboValor.Text = valor.comboText
    Me.txtNumero.Text = valormov.nroComprobante
    Me.dtpEmision.Value = valormov.fechaEmision
    Me.dtpVencimiento.Value = valormov.fechaVencimiento
    Me.txtTitular.Text = valormov.titular
    Me.txtBanco.Text = valormov.banco
    
    Me.cmdAgregarValor.Caption = "Corregir"
    
End Sub

Private Sub txtBanco_GotFocus()

    marcarseleccion Me.txtBanco
    
End Sub

Private Sub txtImporte_GotFocus()

    marcarseleccion Me.txtImporte
    
End Sub

Private Sub txtImporte_LostFocus()

    Me.txtImporte.Text = Format(Val(Me.txtImporte.Text), "0.00")
    
End Sub

Private Sub txtNumero_GotFocus()

    marcarseleccion Me.txtNumero
    
End Sub

Private Sub txtProveedor_GotFocus()

    marcarseleccion Me.txtProveedor
    
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
Dim ctlSrc As New clsCtlSearch

Dim comprobante As New clsDAOComprobante

    ctlSrc.formSearch proveedor, KeyAscii, "Proveedores", db
    
    Me.txtProveedor.Text = proveedor.textSearch
    KeyAscii = 0
    
    Set proveedorglobal = proveedor
    
    comprobante.fillComboTipo Me.cboComprobante, cntTTPagos, db
    
End Sub

Private Sub txtSaldo_GotFocus()

    marcarseleccion Me.txtSaldo
    
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSaldo_LostFocus
    End If
    
End Sub

Private Sub txtSaldo_LostFocus()

    proveedormov.pago = Val(Me.txtSaldo.Text)
    Me.grdFacturas.TextMatrix(row, cntPago) = Format(proveedormov.pago, "0.00")
    Me.txtTotal.Text = Format(totalFacturas, "0.00")
    Me.txtSaldo.Visible = False
    
End Sub

Private Sub txtTitular_GotFocus()

    marcarseleccion Me.txtTitular
    
End Sub

Private Sub txtTotal_GotFocus()

    marcarseleccion Me.txtTotal
    
End Sub

Private Sub txtTotal_LostFocus()

    Me.txtTotal.Text = Format(Me.txtTotal.Text, "0.00")
    
End Sub

