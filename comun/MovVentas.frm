VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MovVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación Electrónica"
   ClientHeight    =   8280
   ClientLeft      =   240
   ClientTop       =   690
   ClientWidth     =   16080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   16080
   Begin VB.ListBox lstFacturas 
      Height          =   4935
      Left            =   12240
      TabIndex        =   72
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtExento 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdPresupuesto 
      Caption         =   "Presupuesto"
      Height          =   375
      Left            =   10080
      TabIndex        =   69
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtTotalDescuento 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   495
      Left            =   8160
      MultiLine       =   -1  'True
      TabIndex        =   64
      Top             =   840
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpActualizacion 
      Height          =   255
      Left            =   2400
      TabIndex        =   62
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   25690113
      CurrentDate     =   40182
   End
   Begin VB.TextBox txtNetoSinDescuento 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   270
      Left            =   4320
      TabIndex        =   59
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtPrecioDescuentoConIVA 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   58
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPrecioVentaConIVA 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   57
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPDescuento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8160
      TabIndex        =   55
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   9120
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtCompra 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtTotalValores 
      Height          =   288
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   6720
      Width           =   1695
   End
   Begin VB.ComboBox cboValor 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   6120
      Width           =   3615
   End
   Begin VB.TextBox txtImporte 
      Height          =   285
      Left            =   5280
      TabIndex        =   15
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   7200
      TabIndex        =   16
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   7200
      TabIndex        =   23
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txtTitular 
      Height          =   285
      Left            =   3360
      TabIndex        =   19
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   4200
      TabIndex        =   20
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   42
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txtIVA 
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   8160
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtArtID 
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtUnitario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboArticulo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   5535
   End
   Begin VB.ComboBox cboComprobante 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtPrefijo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtComprobante 
      Height          =   285
      Left            =   7200
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin Crystal.CrystalReport crpComprobante 
      Left            =   11040
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetalle 
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3413
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker datFecha 
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   25690113
      CurrentDate     =   39075
   End
   Begin MSComCtl2.DTPicker datEmision 
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   25690113
      CurrentDate     =   39074
   End
   Begin MSFlexGridLib.MSFlexGrid grdValores 
      Height          =   1095
      Left            =   1440
      TabIndex        =   22
      Top             =   7080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1931
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker datVto 
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   25690113
      CurrentDate     =   39074
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobantes Emitidos"
      Height          =   195
      Index           =   0
      Left            =   12240
      TabIndex        =   73
      Top             =   600
      Width           =   1650
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Exento"
      Height          =   195
      Index           =   16
      Left            =   480
      TabIndex        =   71
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      Height          =   195
      Index           =   18
      Left            =   4320
      TabIndex        =   68
      Top             =   4680
      Width           =   780
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Neto S/D"
      Height          =   195
      Index           =   17
      Left            =   2400
      TabIndex        =   66
      Top             =   4680
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   65
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ultima Act"
      Height          =   195
      Index           =   11
      Left            =   2400
      TabIndex        =   63
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
      Height          =   195
      Index           =   12
      Left            =   4320
      TabIndex        =   60
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "P.D.s/IVA"
      Height          =   195
      Index           =   14
      Left            =   8160
      TabIndex        =   56
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      Height          =   195
      Index           =   10
      Left            =   9120
      TabIndex        =   54
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "P.Límite"
      Height          =   195
      Index           =   15
      Left            =   9120
      TabIndex        =   53
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   9120
      TabIndex        =   51
      Top             =   6480
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Index           =   22
      Left            =   1440
      TabIndex        =   50
      Top             =   5880
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      Height          =   195
      Index           =   23
      Left            =   5280
      TabIndex        =   49
      Top             =   5880
      Width           =   525
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Index           =   24
      Left            =   7200
      TabIndex        =   48
      Top             =   5880
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Emisión"
      Height          =   195
      Index           =   25
      Left            =   9120
      TabIndex        =   47
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Index           =   26
      Left            =   1440
      TabIndex        =   46
      Top             =   6480
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Titular"
      Height          =   195
      Index           =   27
      Left            =   3360
      TabIndex        =   45
      Top             =   6480
      Width           =   435
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      Height          =   195
      Index           =   28
      Left            =   4200
      TabIndex        =   44
      Top             =   6480
      Width           =   465
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   240
      Top             =   4680
      Width           =   11775
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Index           =   21
      Left            =   10080
      TabIndex        =   39
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Neto"
      Height          =   195
      Index           =   19
      Left            =   6240
      TabIndex        =   38
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "I.V.A."
      Height          =   195
      Index           =   20
      Left            =   8160
      TabIndex        =   37
      Top             =   4680
      Width           =   390
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   " Artículos "
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   34
      Top             =   1320
      Width           =   720
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   240
      Top             =   1440
      Width           =   11775
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Index           =   8
      Left            =   2400
      TabIndex        =   32
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   9
      Left            =   8160
      TabIndex        =   31
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "P.U.s/IVA"
      Height          =   195
      Index           =   13
      Left            =   7200
      TabIndex        =   30
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   28
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Index           =   5
      Left            =   6240
      TabIndex        =   27
      Top             =   600
      Width           =   555
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobante"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   26
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   25
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "MovVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tipocomprobante As String

Private intItem2 As Integer
Private intItem As Integer

Private clientemovimientoId As Long

Private blnDebita As Boolean
Private flagContado As Boolean
Private flagNoBuscar As Boolean
Private blnFlagCarga As Boolean
Private blnTipeandoPrecioDescuento As Boolean
Private blnTipeandoPorcentajeDescuento As Boolean

Private valorMovimiento As New clsDAOValores

Private comprobante As New clsDAOTiposComprob

Private objAr As New clsDAOArticulos

Private clientemov As New clsDAOMovclie

Private cliente As New clsDAOClientes

Private objTMDetArtic As New clsDAOTMDetartic

Private objCgoValor As New clsDAOCgosValores

Private presupuesto As New clsDAOPresupuesto

Private presupuestoDetalle As New clsDAODetpresupuesto

Private vPosItem As Integer
Private vPosCodigo As Integer
Private vPosDescripcion As Integer
Private vPosPUnitario As Integer
Private vPosDescuento As Integer
Private vPosPUcDesc As Integer
Private vPosCantidad As Integer
Private vPosTotal As Integer

Private vDescuento As Currency

Private vAlicuotaiva As Variant

Private Sub fillFacturas()
Dim clienteMovimiento As clsMODClienteMovimiento

Dim clienteMovimientoRep As New clsREPClienteMovimiento

    For Each clienteMovimiento In clienteMovimientoRep.collectionAsociables(cliente.codigo)
        Me.lstFacturas.AddItem clienteMovimiento.tipocomprobante & modConv.formatNumComprobante(clienteMovimiento.puntoventa, clienteMovimiento.numerocomprobante) & " - " & clienteMovimiento.fechacomprobante & "$ " & Format(clienteMovimiento.importe, "0.00")
        Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = clienteMovimiento.clientemovimientoId
    Next
    
End Sub

Private Sub fillTipos()

    vDescuento = 1 - cliente.descuento / 100
    
    tipocomprobante = "B"
    If cliente.posicion = 1 Or cliente.posicion = 4 Then tipocomprobante = "A"
    
    comprobante.fillComboVentasFE Me.cboComprobante, tipocomprobante, 0, gEmpresa.negID, db
    
    If Not modCUIT.validaCUIT(Replace(cliente.cuit, "-", "")) Then MsgBox "ADVERTENCIA: CUIT NO Válido . . ."

End Sub

Private Sub buscaComprob()
Dim intFactor As Integer

Dim objAr As New clsDAOArticulos
Dim objDA As New clsDAODetartic

On Error Resume Next

    blnFlagCarga = False
    
    If Not (Me.cboComprobante.ListIndex >= 0 And IsNumeric(txtPrefijo.Text) And IsNumeric(txtComprobante.Text)) Then Exit Sub
    
    clientemov.findByComprobante gEmpresa.negID, Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex), Val(Me.txtPrefijo.Text), Val(Me.txtComprobante.Text), db
        
    If clientemov.clave > 0 Then
        blnFlagCarga = True
        txtClave.Text = clientemov.clave
        clientemovimientoId = clientemov.clave
        intFactor = IIf(blnDebita, -1, 1)
        txtTotal.Text = Format(a2Decimales(clientemov.importe * intFactor), "0.00")
        txtIVA.Text = Format(a2Decimales(clientemov.montoiva * intFactor), "0.00")
        txtNeto.Text = Format(a2Decimales(clientemov.neto * intFactor), "0.00")
        
        cliente.findByCodigo clientemov.cgoclie, db
        If clientemov.cgoclie <> 0 Then Me.txtCliente.Text = cliente.descripcionBuscar
        
        comprobante.codigo = clientemov.cgocomprob
        comprobante.findByPrimaryKey db
        
        If clientemov.cgocomprob <> 0 Then
            flagNoBuscar = True
            Me.cboComprobante.Text = comprobante.comboText
        End If
        
        If IsDate(clientemov.fechacomprob) Then Me.datFecha.value = clientemov.fechacomprob
        
        intItem = 0
        
        For Each objDA In objDA.collectionByClaveMovClie(Val(Me.txtClave.Text))
            intItem = intItem + 1
            
            Set objTMDetArtic = New clsDAOTMDetartic
            
            objTMDetArtic.hWnd = Me.hWnd
            objTMDetArtic.item = intItem
            objTMDetArtic.cgoarticulo = objDA.cgoartic
            
            objAr.codigo = objDA.cgoartic
            objAr.findByPrimaryKey db
            objTMDetArtic.descripcion = objAr.descripcion
            
            objTMDetArtic.precioventasiniva = CDbl(CLng(objDA.precioventasiniva * 100) / 100)
            objTMDetArtic.cantidad = CDbl(CLng(objDA.cantidad * intFactor * 100) / 100)
            objTMDetArtic.totalsiniva = CDbl(Format(objDA.cantidad * intFactor * objDA.precioventasiniva, "0.00"))
            objTMDetArtic.preciodescuentosiniva = objDA.precioventasiniva
            objTMDetArtic.descuento = objDA.descuento
            
            objTMDetArtic.add
        Next
        
        cmdIngresar.Caption = "Ver Fact"
    Else
        cmdIngresar.Enabled = True
    End If
    
    Call cargaDetalle
    
End Sub

Private Sub borrar()

    Me.cboComprobante.ListIndex = -1
    Me.grdDetalle.Rows = 1
    Me.grdValores.Rows = 1
    Me.txtTotalValores.Text = ""
    txtPrefijo.Text = ""
    txtTotal.Text = ""
    txtNeto.Text = ""
    txtIVA.Text = ""
    txtCantidad.Text = ""
    txtUnitario.Text = ""
    txtDescuento.Text = ""
    txtCompra.Text = ""
    txtPDescuento.Text = ""
    txtItem.Text = "1"
    Me.txtPrefijo.Text = ""
    Me.txtComprobante.Text = ""
    vAlicuotaiva = Null
    
    tipocomprobante = ""
    intItem2 = 1
    blnFlagCarga = False
    
    Me.datFecha.value = Date
    Me.datEmision.value = Date
    Me.datVto.value = Date
    cmdIngresar.Caption = "Ingresar"
    cmdAnular.Visible = False
    cmdGrabar.Caption = "Grabar"
    
    With objTMDetArtic
        .hWnd = Me.hWnd
        
        .deleteAll db
    End With
    
End Sub

Public Sub cargarPendiente(pPreID As Long)
Dim objPar As New clsDAOParametros

    objPar.findLast

    presupuesto.preID = pPreID
    presupuesto.findByPrimaryKey
     
    cliente.findByCodigo presupuesto.cliID, db
    Me.txtCliente.Text = cliente.descripcionBuscar
    
    fillTipos
    
    Me.txtObservaciones.Text = presupuesto.observac
    
    For Each presupuestoDetalle In presupuestoDetalle.collectionByPreID(presupuesto.preID)
        objAr.codigo = presupuestoDetalle.artid
        objAr.findByPrimaryKey db

        If IsNull(vAlicuotaiva) Then
            vAlicuotaiva = objPar.iva1
            If objAr.iva105 Then vAlicuotaiva = objPar.iva2
            If objAr.exento Then vAlicuotaiva = 0
        End If
            
        Set objTMDetArtic = New clsDAOTMDetartic
        
        With objTMDetArtic
            .hWnd = Me.hWnd
            
            .item = intItem2
        End With
        
        intItem = intItem + 1
        intItem2 = intItem2 + 1
        txtItem.Text = intItem2
        
        With objTMDetArtic
            .cgoarticulo = objAr.codigo
            .descripcion = objAr.descripcion & " (" & objAr.codigo & ")"
            .precioventasiniva = presupuestoDetalle.unitsiva
            .precioventaconiva = presupuestoDetalle.unitciva
            .descuento = 0
            .preciodescuentosiniva = presupuestoDetalle.unitsiva
            .preciodescuentoconiva = presupuestoDetalle.unitciva
            .cantidad = presupuestoDetalle.cantidad
            .totalsiniva = presupuestoDetalle.unitsiva * .cantidad
            .totalconiva = presupuestoDetalle.unitciva * .cantidad
            .preciocomprasiniva = objAr.preciocomprasiniva
            
            .save
        End With
    
    Next
    
    txtItem.Text = intItem2
    
    Call cargaDetalle

End Sub

Private Function calcularValores() As Currency
Dim curTotal As Currency

Dim intCiclo As Integer

    curTotal = 0
    
    For intCiclo = 1 To Me.grdValores.Rows - 1
        curTotal = curTotal + Val(Me.grdValores.TextMatrix(intCiclo, 3))
    Next intCiclo
    
    calcularValores = curTotal
    
End Function

Private Sub grabaValores()
Dim intCiclo As Integer

Dim valorMovimiento As clsDAOValores

    For intCiclo = 1 To Me.grdValores.Rows - 1
        Set valorMovimiento = New clsDAOValores
        
        With valorMovimiento
            .codigo = Me.grdValores.TextMatrix(intCiclo, 0)
            .cgocli = cliente.codigo
            .fechaemi = CDate(Me.grdValores.TextMatrix(intCiclo, 6))
            .fechavto = CDate(Me.grdValores.TextMatrix(intCiclo, 7))
            .fechaReg = Me.datFecha.value
            .nroComprob = Val(Me.grdValores.TextMatrix(intCiclo, 2))
            .importe = Val(Me.grdValores.TextMatrix(intCiclo, 3))
            
            objCgoValor.codigo = .codigo
            objCgoValor.findByPrimaryKey
            
            .clavemovv = clientemovimientoId
            .titular = Me.grdValores.TextMatrix(intCiclo, 4)
            .banco = Me.grdValores.TextMatrix(intCiclo, 5)
                    
            .add
        End With
    Next intCiclo
    
End Sub

Private Sub mostrarArticulo()
Dim curPrecioDescuento As Currency
Dim curPorcentajeDescuento As Currency
Dim curAlicuotaIVA As Currency
Dim curPrecioVentaSinIva As Currency
Dim curPrecioVentaConIva As Currency

Dim objAU As New clsDAOArticulosUbic

On Error Resume Next
     
    Me.dtpActualizacion.value = Date
    
    With objAr
        curPrecioVentaSinIva = .precioventasiniva * vDescuento
        curPrecioVentaConIva = .precioventaconiva * vDescuento
        curAlicuotaIVA = 1.21
        If .iva105 Then curAlicuotaIVA = 1.105
        If .exento Then curAlicuotaIVA = 1
        
        If Val(Me.txtPDescuento.Text) = 0 Then Me.txtPDescuento.Text = Format(curPrecioVentaSinIva, "0.00")
        
        If Val(Me.txtPDescuento.Text) > curPrecioVentaSinIva Then
            curPrecioVentaSinIva = Val(Me.txtPDescuento.Text)
            curPrecioVentaConIva = curPrecioVentaSinIva * curAlicuotaIVA
        End If
        
        Me.txtPrecioVentaConIVA.Text = Format(curPrecioVentaConIva, "0.00")
        Me.txtUnitario.Text = Format(curPrecioVentaSinIva, "0.00")
        Me.txtCompra.Text = Format(.preciocomprasiniva, "0.00")
        Me.dtpActualizacion.value = .fechaactualizacion
        
        If Not blnTipeandoPrecioDescuento Then
            curPorcentajeDescuento = Val(Me.txtDescuento.Text)
            curPrecioDescuento = curPrecioVentaSinIva * (1 - (curPorcentajeDescuento / 100))
        Else
            curPrecioDescuento = Val(Me.txtPDescuento.Text)
            curPorcentajeDescuento = 100 * (1 - (curPrecioDescuento / curPrecioVentaSinIva))
        End If
        
        If Not blnTipeandoPrecioDescuento Then
            Me.txtPDescuento.Text = Format(curPrecioDescuento, "0.00")
        Else
            Me.txtDescuento.Text = Format(curPorcentajeDescuento, "0.00")
        End If
        
        Me.txtPrecioDescuentoConIVA.Text = Format(curPrecioVentaConIva * (1 - (curPorcentajeDescuento / 100)), "0.00")
        
        objAU.artid = .codigo
        objAU.findByPrimaryKey
        Me.txtUbicacion.Text = objAU.ubicacion
    End With

End Sub

Private Sub cargaDetalle()
Dim intRow As Integer

Dim curExento As Currency
Dim curNeto As Currency
Dim curNetoSinDescuento As Currency
Dim curTotal As Currency

Dim objTM As New clsDAOTMDetartic

On Error Resume Next

    intRow = 0
    curNeto = 0
    curNetoSinDescuento = 0
    curTotal = 0
    curExento = 0
    
    Me.grdDetalle.Rows = 1
    
    For Each objTM In objTM.collectionByhWnd(Me.hWnd)
        intRow = intRow + 1
        grdDetalle.Rows = intRow + 1
        
        With objTM
            grdDetalle.TextMatrix(intRow, vPosItem) = .item
            grdDetalle.TextMatrix(intRow, vPosCodigo) = .cgoarticulo
            grdDetalle.TextMatrix(intRow, vPosDescripcion) = .descripcion
            grdDetalle.TextMatrix(intRow, vPosPUnitario) = Format(.precioventasiniva, "0.00")
            grdDetalle.TextMatrix(intRow, vPosCantidad) = .cantidad
            grdDetalle.TextMatrix(intRow, vPosTotal) = Format(.totalsiniva, "0.00")
            grdDetalle.TextMatrix(intRow, vPosDescuento) = Format(.descuento, "0.00")
            grdDetalle.TextMatrix(intRow, vPosPUcDesc) = Format(.preciodescuentosiniva, "0.00")
            
            curTotal = curTotal + .totalconiva
            If .exento = 1 Then
                curExento = curExento + .totalconiva
            Else
                curNeto = curNeto + .totalsiniva
                curNetoSinDescuento = curNetoSinDescuento + .precioventasiniva * .cantidad
            End If
        End With
    Next
    
    Me.txtIVA.Text = Format(curTotal - curNeto - curExento, "0.00")
    Me.txtNeto.Text = Format(curNeto, "0.00")
    Me.txtNetoSinDescuento.Text = Format(curNetoSinDescuento, "0.00")
    Me.txtTotalDescuento.Text = Format(curNetoSinDescuento - curNeto, "0.00")
    Me.txtExento.Text = Format(curExento, "0.00")
    Me.txtTotal.Text = Format(curTotal, "0.00")
    
End Sub

Private Sub buscaNroFact()
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    If blnFlagCarga Then Exit Sub
    
    Me.txtComprobante.Text = clientemov.buscaNumeroFactura(Val(Me.txtPrefijo.Text), tipocomprobante, db)
    
    cmdIngresar.Enabled = True
    
End Sub

Private Sub imprime()
Dim factura_service As New clsCtlFactura
Dim impresion_service As New clsCtlImpresion

Dim strReporte As String

    Me.MousePointer = 11
    
    strReporte = IIf(tipocomprobante = "A", "FacturaA", "FacturaB")
    
    With clientemov
        .clave = clientemovimientoId
        .findByPrimaryKey db
    End With
    
    If clientemov.cae = "" Then
        impresion_service.printReport Me.crpComprobante, strReporte, db.sconection, , Array(Array("pClave", clientemovimientoId))
    Else
        factura_service.print_factura clientemovimientoId, clientemov.tipocompro, clientemov.prefijo, clientemov.nroComprob, Me.hWnd
'        If tipoComprobante = "A" Then
'            impresion_service.printReport Me.crpComprobante, "rptFacturaA", db.sconection, Array("sFactura", "sFactura - 01", "sFactura - 02"), Array(Array("pClave", clientemovimientoId))
'        Else
'            impresion_service.printReport Me.crpComprobante, "rptFacturaB", db.sconection, Array("sFactura", "sFactura - 01"), Array(Array("pClave", clientemovimientoId))
'        End If
    End If
    
    Me.MousePointer = 0
    
End Sub

Private Sub grabar()
Dim intFactor As Integer

Dim clientemov As New clsDAOMovclie

Dim objParametros As New clsDAOParametros

Dim strCAE As String
Dim strCAEVenc As String
Dim strBarras As String

Dim lngNroComprob As Long

Dim clientemovimientoIdasociado As Variant

    clientemovimientoIdasociado = Null
    
    objParametros.findLast
    
    If cliente.codigo = 0 Then
        MsgBox "ERROR: Debe seleccionar un Cliente"
        Me.txtCliente.SetFocus
        Exit Sub
    End If
    
    If Me.cboComprobante.ListIndex < 0 Then
        MsgBox "ERROR: Debe seleccionar un Tipo de Comprobante"
        Me.cboComprobante.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtPrefijo.Text) Then
        MsgBox "ERROR: Debe Indicar el prefijo"
        txtPrefijo.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtComprobante.Text) Then
        MsgBox "ERROR: Debe Indicar el Número del Comprobante"
        txtComprobante.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtTotal.Text) Then
        MsgBox "ERROR: Debe Indicar el Importe Total"
        txtTotal.SetFocus
        Exit Sub
    Else
        If Val(txtTotal.Text) = 0 Then
            MsgBox "ERROR: Imposible facturar valor cero"
            txtArtID.SetFocus
            Exit Sub
        End If
    End If
    
    If Not IsNumeric(txtNeto.Text) Then txtNeto.Text = 0
    If Not IsNumeric(Me.txtNetoSinDescuento.Text) Then Me.txtNetoSinDescuento.Text = 0
    If Not IsNumeric(txtIVA.Text) Then txtIVA.Text = 0
    If Not IsNumeric(txtCantidad.Text) Then txtCantidad.Text = 0
    If Not IsNumeric(txtUnitario.Text) Then txtUnitario.Text = 0
    
    intFactor = IIf(blnDebita, 1, -1)
    
    strCAE = ""
    strCAEVenc = ""
    If comprobante.factelect <> 0 Then
        If Me.lstFacturas.ListIndex > -1 Then clientemovimientoIdasociado = Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex)
        strCAE = modFEv1.cae(comprobante.codigo, cliente.codigo, Val(Me.txtTotal.Text), Val(Me.txtExento.Text), Val(Me.txtNeto.Text), Val(Me.txtIVA.Text), objParametros.feproduccion, vAlicuotaiva, db, lngNroComprob, strCAEVenc, strBarras, clientemovimientoIdasociado)
        If strCAE = "" Then Exit Sub
        Me.txtComprobante.Text = lngNroComprob
    End If
    
    With clientemov
        .negID = gEmpresa.negID
        .cgoclie = cliente.codigo
        .cgocomprob = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
        .fechacomprob = Me.datFecha.value
        .fechavenc = Me.datFecha.value
        .importe = CDbl(Me.txtTotal.Text) * intFactor
        .prefijo = Val(Me.txtPrefijo.Text)
        .nroComprob = Val(Me.txtComprobante.Text)
        .montoiva = CDbl(Me.txtIVA.Text) * intFactor
        .montoexento = CDbl(Me.txtExento.Text) * intFactor
        .neto = CDbl(txtNeto.Text) * intFactor
        .netosindescuento = CDbl(Me.txtNetoSinDescuento.Text) * intFactor
        .tipocompro = tipocomprobante
        .iva = vAlicuotaiva
        .observaciones = Me.txtObservaciones.Text
        .cae = strCAE
        .caevenc = strCAEVenc
        .barras = strBarras
        
        If flagContado Then .cancelado = .importe
    
        .add db
        
        clientemovimientoId = .clave
    End With
    
    Call grabaDetArtic
    Call grabaValores
            
    With presupuesto
        If .preID > 0 Then
            .clavemovclie = clientemovimientoId
            .update
        End If
    End With
    
    Call imprime
    
    Call cmdCancelar_Click
    
End Sub

Private Sub grabaDetArtic()
Dim intFactor As Integer

Dim objDetArtic As New clsDAODetartic

Dim objTM As New clsDAOTMDetartic

    intFactor = IIf(blnDebita, -1, 1)
    
    For Each objTM In objTM.collectionByhWnd(Me.hWnd)
        With objDetArtic
            .clavemovclie = clientemovimientoId
            .cgocomprob = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
            .item = objTM.item
            .cgoartic = objTM.cgoarticulo
            .descripcion = objTM.descripcion
            .cantidad = objTM.cantidad * intFactor
            .descuento = objTM.descuento
            .preciocomprasiniva = objTM.preciocomprasiniva
            .precioventaconiva = objTM.precioventaconiva
            .precioventasiniva = objTM.precioventasiniva
            .preciodescuentoconiva = objTM.preciodescuentoconiva
            .preciodescuentosiniva = objTM.preciodescuentosiniva
            
            .add
        End With
    Next

End Sub

Private Sub cboArticulo_Click()

    If Me.cboArticulo.ListIndex < 0 Then Exit Sub
    
    Me.txtPDescuento.Text = ""
    
    With objAr
        .findByClave Me.cboArticulo.ItemData(Me.cboArticulo.ListIndex), db
        
        Me.txtArtID.Text = .codigo
        
        Call mostrarArticulo
    End With
    
End Sub

Private Sub cboArticulo_GotFocus()

    SendMessageLong Me.cboArticulo.hWnd, &H14F, True, 0

End Sub

Private Sub cboArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
Static intPosicion As Integer

Dim intContador As Integer
    
Dim colArticulos As New Collection

    If intPosicion = cboArticulo.SelStart Then
        intPosicion = intPosicion - 2
    Else
        intPosicion = cboArticulo.SelStart
    End If
    
    If intPosicion > 3 Then
        intContador = 0
        Set colArticulos = objAr.collectionByDescripcion(Mid(Me.cboArticulo.Text, 1, intPosicion))
        
        Me.cboArticulo.Clear
        
        For Each objAr In colArticulos
            intContador = intContador + 1
            
            With objAr
                If intContador = 1 Then
                    Me.txtArtID.Text = .codigo
                    Me.cboArticulo.Text = .descripcion & " (" & .codigo & ")"
                End If
                
                Me.cboArticulo.AddItem .descripcion & " (" & .codigo & ")"
                Me.cboArticulo.ItemData(Me.cboArticulo.NewIndex) = .clave
                    
            End With
        Next
        
        cboArticulo.SelStart = intPosicion
        cboArticulo.SelLength = Len(cboArticulo.Text)
    End If

End Sub

Private Sub cboArticulo_LostFocus()

    If Me.cboArticulo.ListIndex < 0 Then Exit Sub
    
    With objAr
        .findByClave Me.cboArticulo.ItemData(Me.cboArticulo.ListIndex), db
        
        Me.txtArtID.Text = .codigo
        
        Call mostrarArticulo
    End With

End Sub

Private Sub cboComprobante_Click()

    Me.lstFacturas.Clear
    If cboComprobante.ListIndex < 0 Then Exit Sub
    
    comprobante.codigo = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
    comprobante.findByPrimaryKey db
    tipocomprobante = comprobante.tipocomprob
    blnDebita = IIf(comprobante.debita = 0, False, True)
    flagContado = IIf(comprobante.contado = 0, False, True)
    If comprobante.puntovta > 0 Then Me.txtPrefijo.Text = comprobante.puntovta
    
    Me.Height = Me.cmdSalir.Top + Me.cmdSalir.Height + 450 + IIf(flagContado, 2400, 0)
    
    If comprobante.asociado Then fillFacturas
    
    If Not flagNoBuscar Then Call buscaNroFact
    
End Sub

Private Sub cmdAgregar_Click()

    If Me.cboValor.ListIndex < 0 Then Exit Sub
    
    If Val(Me.txtImporte.Text) = 0 Then Exit Sub
    
    Me.cmdAgregar.Enabled = False
    
    With Me.grdValores
        .Rows = .Rows + 1
        
        .TextMatrix(.Rows - 1, 0) = Me.cboValor.ItemData(Me.cboValor.ListIndex)
        .TextMatrix(.Rows - 1, 1) = Me.cboValor.Text
        .TextMatrix(.Rows - 1, 2) = Me.txtNumero.Text
        .TextMatrix(.Rows - 1, 3) = Format(Val(Me.txtImporte.Text), "0.00")
        .TextMatrix(.Rows - 1, 4) = Me.txtTitular.Text
        .TextMatrix(.Rows - 1, 5) = Me.txtBanco.Text
        .TextMatrix(.Rows - 1, 6) = Me.datEmision.value
        .TextMatrix(.Rows - 1, 7) = Me.datVto.value
    End With
    
    Me.txtTotalValores.Text = Format(calcularValores, "0.00")
    
    Me.cmdAgregar.Enabled = True

End Sub

Private Sub cmdEliminar_Click()

    If Me.grdValores.row = 0 Then Exit Sub
    
    Me.cmdEliminar.Enabled = False
    
    If Me.grdValores.Rows = 2 Then
        Me.grdValores.Rows = 1
    Else
        Me.grdValores.RemoveItem Me.grdValores.row
    End If
    
    Me.txtTotalValores.Text = Format(calcularValores, "0.00")
    
    Me.cmdEliminar.Enabled = True

End Sub

Private Sub cmdIngresar_Click()

    If comprobante.contado <> 0 And Val(Me.txtTotalValores.Text) <> Val(Me.txtTotal.Text) Then
        MsgBox "No coinciden los VALORES"
        Exit Sub
    End If
    
    cmdIngresar.Enabled = False
    
    If Me.cmdIngresar.Caption = "Ingresar" Then
        Call grabar
    Else
        Call imprime
    End If
        
    Me.txtCantidad.SetFocus
    
End Sub

Private Sub cmdPresupuesto_Click()
Dim lngPreID As Long

    borrar

    lngPreID = 0

    Load frmPresupPendiente
           
    frmPresupPendiente.Show vbModal
    
    If Not IsNull(frmPresupPendiente.preID) Then lngPreID = frmPresupPendiente.preID
    
    Unload frmPresupPendiente
    
    If lngPreID > 0 Then cargarPendiente lngPreID
    
End Sub

Private Sub cmdSalir_LostFocus()

    Me.txtCliente.SetFocus
    
End Sub

Private Sub cmdCancelar_Click()

    txtPrefijo.Text = ""
    txtComprobante.Text = ""
    txtTotal.Text = ""
    txtNeto.Text = ""
    txtIVA.Text = ""
    txtItem.Text = "1"
    txtObservaciones.Text = ""
    flagNoBuscar = False
    presupuesto.preID = 0
    
    Call borrar
    
    cmdIngresar.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cmdGrabar_Click()
Dim objAr As New clsDAOArticulos

Dim objParametro As New clsDAOParametros

Dim curAlicuotaIVA As Currency

    objParametro.findLast

    If Not IsNumeric(txtUnitario.Text) Then
        MsgBox "ERROR: Debe Indicar Precio"
        txtUnitario.SetFocus
        Exit Sub
    End If
    
    If txtArtID.Text = "" Then
        MsgBox "ERROR: Debe Indicar Articulo"
        txtArtID.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtCantidad.Text) Then
        MsgBox "ERROR: Debe Indicar cantidad"
        txtCantidad.SetFocus
        Exit Sub
    End If
    
    objAr.codigo = Trim(Me.txtArtID.Text)
    objAr.findByPrimaryKey db
    
    If objAr.clave = 0 Then
        MsgBox "ERROR: Artículo NO REGISTRADO"
        Me.txtArtID.SetFocus
        Exit Sub
    End If
    
    curAlicuotaIVA = objParametro.iva1
    If objAr.iva105 Then curAlicuotaIVA = objParametro.iva2
    If objAr.exento Then curAlicuotaIVA = 0
    
    If IsNull(vAlicuotaiva) Then
        vAlicuotaiva = objParametro.iva1
        If objAr.iva105 Then vAlicuotaiva = objParametro.iva2
        If objAr.exento Then vAlicuotaiva = 0
    End If
    
    If cmdGrabar.Caption = "Grabar" Then
        Set objTMDetArtic = New clsDAOTMDetartic
        
        With objTMDetArtic
            .hWnd = Me.hWnd
            .item = intItem2
        End With
        
        intItem = intItem + 1
        intItem2 = intItem2 + 1
        txtItem.Text = intItem2
    Else
        With objTMDetArtic
            .hWnd = Me.hWnd
            .item = Val(Me.txtItem.Text)
            
            .findByItem db
        End With
        cmdAnular.Visible = False
        cmdGrabar.Caption = "Grabar"
    End If
    
    With objTMDetArtic
        .cgoarticulo = txtArtID.Text
        .descripcion = objAr.descripcion
        .precioventasiniva = Me.txtUnitario.Text
        .precioventaconiva = Me.txtPrecioVentaConIVA.Text
        .descuento = Val(Me.txtDescuento.Text)
        .preciodescuentosiniva = Me.txtPDescuento.Text
        .preciodescuentoconiva = Me.txtPrecioDescuentoConIVA.Text
        .cantidad = txtCantidad.Text
        .totalsiniva = .preciodescuentosiniva * Val(txtCantidad.Text)
        .totalconiva = .preciodescuentoconiva * Val(txtCantidad.Text)
        .preciocomprasiniva = Me.txtCompra.Text
        .exento = objAr.exento
        
        .save
    End With
    
    cmdGrabar.Caption = "Grabar"
    txtItem.Text = intItem2
    
    Call cargaDetalle
    
    txtArtID.SetFocus
    
End Sub

Private Sub cmdAnular_Click()

    With objTMDetArtic
        .hWnd = Me.hWnd
        .item = Val(Me.txtItem.Text)
        
        .eliminarItem db
        .renumerarItem db
    End With

    cmdAnular.Visible = False
    cmdGrabar.Caption = "Grabar"
    
    Call cargaDetalle
    
    txtItem.Text = Me.grdDetalle.Rows + 1
    
End Sub

Private Sub Form_Load()
Dim strTitulos As Variant
Dim intAnchos As Variant

    blnTipeandoPorcentajeDescuento = False
    blnTipeandoPrecioDescuento = False

    strTitulos = Array("", "Item", "Codigo", "Descripcion", "P.Unitario", "Descuento", "P.U.c/Desc", "Cantidad", "Total")
    intAnchos = Array(0, 500, 1200, 4100, 1000, 1000, 1000, 1000, 1000)
    vPosItem = 1
    vPosCodigo = 2
    vPosDescripcion = 3
    vPosPUnitario = 4
    vPosDescuento = 5
    vPosPUcDesc = 6
    vPosCantidad = 7
    vPosTotal = 8
    
    modGrid.makeGrid Me.grdDetalle, strTitulos, intAnchos, 1, 1, flexSelectionByRow
    
    strTitulos = Array("Codigo", "Concepto", "Numero", "Importe", "Titular", "Banco", "Emision", "Vencimiento")
    intAnchos = Array(0, 3000, 1000, 1000, 1000, 1000, 1000, 1000)
    
    modGrid.makeGrid Me.grdValores, strTitulos, intAnchos, 0, 1, flexSelectionByRow
    
    With Me.grdDetalle
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns
    End With
    
    objCgoValor.fillCombo Me.cboValor
        
    Call borrar
    
End Sub

Private Sub datFecha_LostFocus()

    Me.datEmision.value = Me.datFecha.value
    Me.datVto.value = Me.datFecha.value
    
End Sub

Private Sub grdDetalle_DblClick()
Dim intRow As Integer

    With Me.grdDetalle
        intRow = .row
        Me.txtItem.Text = .TextMatrix(intRow, vPosItem)
        Me.txtArtID.Text = .TextMatrix(intRow, vPosCodigo)
        Me.cboArticulo.Text = .TextMatrix(intRow, vPosDescripcion)
        Me.txtUnitario.Text = .TextMatrix(intRow, vPosPUnitario)
        Me.txtCantidad.Text = .TextMatrix(intRow, vPosCantidad)
        Me.txtDescuento.Text = .TextMatrix(intRow, vPosDescuento)
        Me.txtPDescuento.Text = .TextMatrix(intRow, vPosPUcDesc)
    End With
    
    Me.cmdGrabar.Caption = "Regrabar"
    Me.cmdAnular.Visible = True
    
End Sub

Private Sub txtArtID_GotFocus()

    blnTipeandoPorcentajeDescuento = False
    blnTipeandoPrecioDescuento = False

    marcarseleccion Me.txtArtID
    
End Sub

Private Sub txtCantidad_LostFocus()

    Call mostrarArticulo
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    
    cliente.formBuscar frmBuscar, cliente, KeyAscii, "Clientes"
    
    Me.txtCliente.Text = cliente.descripcionBuscar
    KeyAscii = 0
    
    fillTipos
    Me.lstFacturas.Clear

End Sub

Private Sub txtComprobante_GotFocus()

    marcarseleccion Me.txtComprobante
    
End Sub

Private Sub txtDescuento_Change()

    If blnTipeandoPorcentajeDescuento Then Call mostrarArticulo
    
End Sub

Private Sub txtDescuento_GotFocus()

    marcarseleccion Me.txtDescuento
    
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)

    blnTipeandoPorcentajeDescuento = True
    blnTipeandoPrecioDescuento = False
    
    If KeyAscii <> 13 Then Exit Sub

    Me.cmdGrabar.SetFocus
    
End Sub

Private Sub txtIVA_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtIVA_LostFocus
    End If
    
End Sub

Private Sub txtIVA_LostFocus()

On Error Resume Next

    Me.txtNeto.Text = Format(CDbl(Me.txtTotal.Text) - CDbl(Me.txtIVA.Text), "0.00")

End Sub

Private Sub txtObservaciones_GotFocus()

    marcarseleccion Me.txtObservaciones
    
End Sub

Private Sub txtPDescuento_Change()

    If blnTipeandoPrecioDescuento Then Call mostrarArticulo
    
End Sub

Private Sub txtPDescuento_GotFocus()

    marcarseleccion Me.txtPDescuento
    
End Sub

Private Sub txtPDescuento_KeyPress(KeyAscii As Integer)

    blnTipeandoPrecioDescuento = True
    blnTipeandoPorcentajeDescuento = False
    
End Sub

Private Sub txtPrefijo_LostFocus()

    Call buscaNroFact

End Sub

Private Sub txtComprobante_LostFocus()

    Call buscaComprob

End Sub

Private Sub txtCantidad_GotFocus()

    blnTipeandoPorcentajeDescuento = False
    blnTipeandoPrecioDescuento = False

    marcarseleccion Me.txtCantidad

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    
    Me.txtDescuento.SetFocus

End Sub

Private Sub txtArtID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    If txtArtID.Text <> "" Then
        txtCantidad.Text = ""
        txtCantidad.SetFocus
    Else
        cmdIngresar.SetFocus
    End If

End Sub

Private Sub txtArtID_LostFocus()
Dim objArticulosAlias As New clsDAOArticulosAlias

    If txtArtID.Text = "" Then Exit Sub
    
    Me.txtPDescuento.Text = ""
    
    With objAr
        .codigo = Me.txtArtID.Text
        .findByPrimaryKey db
        
        If .clave = 0 Then
            objArticulosAlias.findByAlias Me.txtArtID.Text, db
            
            If objArticulosAlias.clave <> 0 Then
                .codigo = objArticulosAlias.artid
                .findByPrimaryKey db
            End If
            Me.txtArtID.Text = .codigo
        End If
        
        If .descripcion = "" Then
            Me.txtArtID.Text = ""
            Me.cboArticulo.ListIndex = -1
            Exit Sub
        End If
        
        If Me.cboArticulo.Text <> .descripcion & " (" & .codigo & ")" Then cboArticulo.Text = .descripcion & " (" & .codigo & ")"
            
        Call mostrarArticulo
    End With
    
End Sub

Private Sub txtUnitario_GotFocus()

    marcarseleccion Me.txtUnitario
    
End Sub

