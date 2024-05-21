VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProveedorAnul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anula Movimientos Proveedores"
   ClientHeight    =   7470
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11790
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   255
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
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   5040
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid grdComprobantes 
      Height          =   3615
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   78381057
      CurrentDate     =   39072
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   78381057
      CurrentDate     =   39072
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetalle 
      Height          =   2055
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   3
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Movimientos Proveedor"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmProveedorAnul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private proveedor As New clsDAOProveedor

Private comprobantes As Collection
Private proveedormovs As Collection
Private valores As Collection

Private Sub cleanform()

    Set proveedormovs = Nothing
    Set proveedor = New clsDAOProveedor
    If proveedorglobal.proveedorID > 0 Then Set proveedor = proveedorglobal
    Me.txtProveedor.Text = proveedor.textSearch
        
    Me.dtpDesde.Value = Date
    Me.dtpHasta.Value = Date
    
    Me.grdComprobantes.Rows = 1
    Me.grdDetalle.Rows = 1
    
End Sub

Private Sub fillComprobantes()
Dim comprobante As clsDAOComprobante
Dim proveedormov As New clsDAOProveedorMov
    
    Me.grdComprobantes.Rows = 1
    
    Set proveedormovs = proveedormov.collectionPeriodoByProveedorID(proveedor.proveedorID, Me.dtpDesde.Value, Me.dtpHasta.Value, db)
    
    For Each proveedormov In proveedormovs
        Set comprobante = New clsDAOComprobante
        If modCollection.collectionExistElement(comprobantes, "k." & proveedormov.comprobanteID) Then Set comprobante = comprobantes("k." & proveedormov.comprobanteID)
        Me.grdComprobantes.AddItem modGrid.array2itemGrid(Array(comprobante.descripcion, modConv.formatNumComprobante(proveedormov.prefijo, proveedormov.nroComprobante), proveedormov.fechaComprobante, Format(proveedormov.total, "0.00"), Format(proveedormov.totalCancelado, "0.00"), Format(proveedormov.total - proveedormov.totalCancelado, "0.00"), Format(proveedormov.totalCancelado, "0.00")))
        Me.grdComprobantes.RowData(Me.grdComprobantes.Rows - 1) = proveedormov.proveedormovimientoID
    Next
    
End Sub

Private Sub fillComprobantesDetalle(proveedormovimientoID As Long, aplicacion As Integer)
Dim comprobante As clsDAOComprobante
Dim proveedormov As New clsDAOProveedorMov

Dim proveedormovs As Collection
    
    If aplicacion = 0 Then
        Set proveedormovs = proveedormov.collectionByProveedorMovimientoIDDeuda(proveedormovimientoID, db)
    Else
        Set proveedormovs = proveedormov.collectionByProveedorMovimientoIDPago(proveedormovimientoID, db)
    End If
    
    Me.grdDetalle.Rows = 1
    Me.grdDetalle.Redraw = False
    For Each proveedormov In proveedormovs
        Set comprobante = New clsDAOComprobante
        If modCollection.collectionExistElement(comprobantes, "k." & proveedormov.comprobanteID) Then Set comprobante = comprobantes("k." & proveedormov.comprobanteID)
        Me.grdDetalle.AddItem modGrid.array2itemGrid(Array(comprobante.descripcion, modConv.formatNumComprobante(proveedormov.prefijo, proveedormov.nroComprobante), proveedormov.fechaComprobante, Format(proveedormov.total, "0.00"), Format(proveedormov.totalCancelado, "0.00"), Format(proveedormov.total - proveedormov.totalCancelado, "0.00"), Format(proveedormov.totalCancelado, "0.00")))
    Next
    Me.grdDetalle.Redraw = True
    
End Sub

Private Sub fillValoresDetalle(proveedormovimientoID As Long)
Dim valormov As New clsDAOValorMov
Dim valor As New clsDAOValor

    Me.grdDetalle.Rows = 1
    Me.grdDetalle.Redraw = False
    For Each valormov In valormov.collectionByProveedorMovimientoID(proveedormovimientoID, db)
        Set valor = valores("k." & valormov.valorID)
        Me.grdDetalle.AddItem modGrid.array2itemGrid(Array(valor.comboText, Format(valormov.importe, "0.00"), valormov.nroComprobante, valormov.fechaEmision, valormov.fechaVencimiento, valormov.titular, valormov.banco))
    Next
    Me.grdDetalle.Redraw = True

End Sub

Private Sub cmdAnular_Click()
Dim comprobante As New clsDAOComprobante
Dim proveedormov As clsDAOProveedorMov

Dim ctlCmp As New clsCtlCompras

    If Me.grdComprobantes.Rows = 1 Then Exit Sub

    Set proveedormov = proveedormovs("k." & Me.grdComprobantes.RowData(Me.grdComprobantes.row))
    If modCollection.collectionExistElement(comprobantes, "k." & proveedormov.comprobanteID) Then Set comprobante = comprobantes("k." & proveedormov.comprobanteID)
    
    If MsgBox("Está SEGURO?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Está Realmente SEGURO?", vbYesNo) = vbNo Then Exit Sub
    
    If comprobante.aplicable = 1 And proveedormov.totalCancelado <> 0 Then
        MsgBox "ERROR: Debe Anular Pagos APLICADOS"
        Exit Sub
    End If
    
    Me.cmdAnular.Enabled = False
    Me.MousePointer = 11
    
    If Not ctlCmp.deleteComprobante(proveedormov, db) Then MsgBox "ERROR: No se pudo ELIMINAR Comprobante"
    
    Me.MousePointer = 0
    Me.cmdAnular.Enabled = True

    fillComprobantes
    
End Sub

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.Value = Me.dtpDesde.Value
    
End Sub

Private Sub cmdImprimir_Click()
Dim comprobante As New clsDAOComprobante
Dim proveedormov As clsDAOProveedorMov

Dim ctlImp As New clsCtlImpresion

    If Me.grdComprobantes.Rows = 1 Then Exit Sub
    
    Set proveedormov = proveedormovs("k." & Me.grdComprobantes.RowData(Me.grdComprobantes.row))
    If modCollection.collectionExistElement(comprobantes, "k." & proveedormov.comprobanteID) Then Set comprobante = comprobantes("k." & proveedormov.comprobanteID)
    
    Me.cmdImprimir.Enabled = False
    Me.MousePointer = 11
    
    If comprobante.letraComprobante <> "" Then
        ' Imprime Factura
        ctlImp.printReport Me.crpConsulta, "rptCompraProv", db.sconection, , Array(Array("proveedormovimientoID", proveedormov.proveedormovimientoID))
    Else
        ' Imprime Pago o Aplicacion
        ctlImp.printReport Me.crpConsulta, "rptPagoProv", db.sconection, Array("FactPagadas", "Valores"), Array(Array("proveedormovimientoID", proveedormov.proveedormovimientoID))
    End If
    
    Me.MousePointer = 0
    Me.cmdImprimir.Enabled = True

End Sub

Private Sub cmdRevisar_Click()

    fillComprobantes
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()
Dim comprobante As New clsDAOComprobante
Dim valor As New clsDAOValor

    Set comprobantes = comprobante.collectionAll(db)
    Set valores = valor.collectionAll(db)
    
End Sub

Private Sub Form_Load()
    
    modGrid.makeGrid2 Me.grdComprobantes, Array(Array("Comprobante", 4400), Array("Numero", 1500), Array("Fecha", 1000), Array("Total", 1000), Array("A Cta", 1000), Array("Saldo", 1000), Array("Pago", 1000)), 0, 1, flexSelectionByRow
    modGrid.makeGrid2 Me.grdDetalle, Array(Array("Comprobante", 4400), Array("Numero", 1500), Array("Fecha", 1000), Array("Total", 1000), Array("A Cta", 1000), Array("Saldo", 1000), Array("Pago", 1000)), 0, 1, flexSelectionFree
    
    cleanform
    
End Sub

Private Sub grdComprobantes_Click()
Dim proveedormov As clsDAOProveedorMov
Dim comprobante As New clsDAOComprobante
    
    If Me.grdComprobantes.Rows = 1 Then Exit Sub
    
    Set proveedormov = proveedormovs("k." & Me.grdComprobantes.RowData(Me.grdComprobantes.row))
    If modCollection.collectionExistElement(comprobantes, "k." & proveedormov.comprobanteID) Then Set comprobante = comprobantes("k." & proveedormov.comprobanteID)
    
    If comprobante.letraComprobante <> "" Or comprobante.aplicacion <> 0 Then
        modGrid.makeGrid2 Me.grdDetalle, Array(Array("Comprobante", 4400), Array("Numero", 1500), Array("Fecha", 1000), Array("Total", 1000), Array("A Cta", 1000), Array("Saldo", 1000), Array("Pago", 1000)), 0, 1, flexSelectionFree
        fillComprobantesDetalle proveedormov.proveedormovimientoID, comprobante.aplicacion
    Else
        modGrid.makeGrid2 Me.grdDetalle, Array(Array("Concepto", 4400), Array("Importe", 1500), Array("Numero", 1000), Array("Emision", 1000), Array("Vencimiento", 1000), Array("Titular", 1000), Array("Banco", 1000)), 0, 1, flexSelectionByRow
        fillValoresDetalle proveedormov.proveedormovimientoID
    End If

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
    
    Me.grdComprobantes.Rows = 1
    Me.grdDetalle.Rows = 1
    
End Sub


