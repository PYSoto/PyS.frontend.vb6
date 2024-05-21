VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmArticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos"
   ClientHeight    =   7590
   ClientLeft      =   2235
   ClientTop       =   1170
   ClientWidth     =   13665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   13665
   Begin VB.TextBox txtMovs 
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminarArt 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   6000
      TabIndex        =   56
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtProvPrecioCompra 
      Height          =   285
      Left            =   4920
      TabIndex        =   40
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtUnitCompraAnterior 
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtUtilidadLta 
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtUtilidadVta 
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtCatalogo 
      Height          =   285
      Left            =   6000
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtMarca 
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox lstArticuloAlternativo 
      Height          =   1425
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   48
      Top             =   6000
      Width           =   5535
   End
   Begin VB.TextBox txtAlternativo 
      Height          =   285
      Left            =   7920
      TabIndex        =   47
      Top             =   5640
      Width           =   5535
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   7920
      TabIndex        =   46
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdEliminarAlias 
      Caption         =   "Eliminar Alias"
      Height          =   255
      Left            =   6000
      TabIndex        =   45
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgregarAlias 
      Caption         =   "Agregar Alias"
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtAlias 
      Height          =   285
      Left            =   240
      TabIndex        =   41
      Top             =   5400
      Width           =   855
   End
   Begin VB.ComboBox cboProveedorAlias 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   4080
      TabIndex        =   37
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   6000
      TabIndex        =   36
      Top             =   7200
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdAlternativo 
      Height          =   1215
      Left            =   240
      TabIndex        =   35
      Top             =   6000
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2143
      _Version        =   393216
   End
   Begin VB.ComboBox cboProveedor 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   2640
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker dtpActualizacion 
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   77135873
      CurrentDate     =   39930
   End
   Begin VB.TextBox txtDescuento 
      Height          =   285
      Left            =   6000
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtOrigen 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtModeloCamion 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtPrecioListaConIVA 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtPrecioListaSinIVA 
      Height          =   285
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ListBox lstArticulos 
      Height          =   4740
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtPrecioVentaSinIVA 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtUnitVenta 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtUnitCompra 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkIvaExento 
      Caption         =   "Iva Exento"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chkIva105 
      Caption         =   "IVA 10,5"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   3720
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdAlias 
      Height          =   1215
      Left            =   240
      TabIndex        =   42
      Top             =   4200
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2143
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpUltimaCompra 
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   77135873
      CurrentDate     =   39930
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "U"
      Height          =   195
      Left            =   6840
      TabIndex        =   55
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "U"
      Height          =   195
      Left            =   6840
      TabIndex        =   54
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Libro/Catálogo"
      Height          =   195
      Left            =   4800
      TabIndex        =   51
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      Height          =   195
      Left            =   1560
      TabIndex        =   50
      Top             =   3000
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Artículo Alternativo"
      Height          =   195
      Left            =   7920
      TabIndex        =   49
      Top             =   5400
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Alias"
      Height          =   195
      Left            =   240
      TabIndex        =   43
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Alternativo"
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   5760
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   1320
      TabIndex        =   34
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
      Height          =   195
      Left            =   1320
      TabIndex        =   33
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      Height          =   195
      Left            =   5040
      TabIndex        =   32
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Origen"
      Height          =   195
      Left            =   1560
      TabIndex        =   31
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Modelo Camion"
      Height          =   195
      Left            =   960
      TabIndex        =   30
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Precio Unit. Lista c/IVA"
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Lta s/IVA"
      Height          =   195
      Left            =   4200
      TabIndex        =   28
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Vta s/IVA"
      Height          =   195
      Left            =   4200
      TabIndex        =   26
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Left            =   1560
      TabIndex        =   25
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   1200
      TabIndex        =   24
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Precio Unit. Venta c/IVA"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Precio Unit.Compra s/IVA"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnCargando As Boolean

Private objArticulo As New clsDAOArticulos

Private objProveedor As New clsDAOProveedor

Private objArticulosAlternativo As New clsDAOArticulosAlter

Private objArticulosAlias As New clsDAOArticulosAlias

Private vEscribiendoSinIVA As Boolean
Private vEscribiendoConIVA As Boolean

Private vPrecioCompraOriginal As Currency

Private Sub cboProveedor_Click()

    If Me.cboProveedor.ListIndex < 0 Then Exit Sub
    
    If blnCargando Then Exit Sub
    
    objArticulo.prvID = Me.cboProveedor.ItemData(Me.cboProveedor.ListIndex)
    
End Sub

Private Sub chkIva105_Click()

    If blnCargando Then Exit Sub

    objArticulo.iva105 = Me.chkIva105.Value
    
End Sub

Private Sub chkIvaExento_Click()

    If blnCargando Then Exit Sub

    objArticulo.exento = Me.chkIvaExento.Value
    
End Sub

Private Sub cmdAgregar_Click()
Dim objArticuloLocal As New clsDAOArticulos

Dim objAAP As New clsVDAOAlterArtProv

    If Me.lstArticuloAlternativo.ListIndex < 0 Then
        MsgBox "ERROR: Debe Elegir UN Artículo Alternativo"
        Exit Sub
    End If
    
    If Me.cmdGrabar.Caption = "&Grabar" Then
        MsgBox "ERROR: Artículo NO Registrado Todavía"
        Me.cmdGrabar.SetFocus
        Exit Sub
    End If
    
    objArticuloLocal.findByClave Me.lstArticuloAlternativo.ItemData(Me.lstArticuloAlternativo.ListIndex), db
    
    If Me.txtCodigo.Text = objArticuloLocal.codigo Then
        MsgBox "ERROR: Códigos IGUALES"
        Exit Sub
    End If
    
    Set objArticulosAlternativo = New clsDAOArticulosAlter
    
    With objArticulosAlternativo
        .artid = Me.txtCodigo.Text
        .artidalternativo = objArticuloLocal.codigo
        
        .add
        
        gArtMirror.saveArticuloAlter objArticulosAlternativo, db
        
        objAAP.llenarGrilla Me.grdAlternativo, .artid, db
    End With
    
End Sub

Private Sub cmdAgregarAlias_Click()
Dim objAPr As New clsVDAOAliasProv

    If Me.cboProveedorAlias.ListIndex < 0 Then Exit Sub
    
    If Trim(Me.txtAlias.Text) = "" Then Exit Sub
    
    If Me.cmdGrabar.Caption = "&Grabar" Then
        MsgBox "ERROR: Artículo NO Registrado Todavía"
        Me.cmdGrabar.SetFocus
        Exit Sub
    End If
    
    With objArticulosAlias
        .artid = Me.txtCodigo.Text
        .alias = Trim(Me.txtAlias.Text)
        .prvID = Me.cboProveedorAlias.ItemData(Me.cboProveedorAlias.ListIndex)
        .preciocompra = Val(Me.txtProvPrecioCompra.Text)
        
        .add
        
        gArtMirror.saveArticuloAlias objArticulosAlias, db
        
        objAPr.llenarGrilla Me.grdAlias, .artid, db
    End With
    
End Sub

Private Sub cmdEliminar_Click()
Dim objAAP As New clsVDAOAlterArtProv
    
    If Me.grdAlternativo.row < 1 Then Exit Sub
    
    With objArticulosAlternativo
        .clave = Me.grdAlternativo.TextMatrix(Me.grdAlternativo.row, 0)
        
        .delete
        
        gArtMirror.deleteArticuloAlter objArticulosAlternativo, db
    End With
    
    objAAP.llenarGrilla Me.grdAlternativo, Me.txtCodigo.Text, db
    
End Sub

Private Sub cmdEliminarAlias_Click()
Dim objAPr As New clsVDAOAliasProv

    If Me.grdAlias.row < 1 Then Exit Sub
    
    With objArticulosAlias
        .clave = Me.grdAlias.TextMatrix(Me.grdAlias.row, 0)
        
        .delete
        
        gArtMirror.deleteArticuloAlias objArticulosAlias, db
        
        objAPr.llenarGrilla Me.grdAlias, Me.txtCodigo.Text, db
    End With
    
End Sub

Private Sub cmdEliminarArt_Click()

    If Trim(Me.txtCodigo.Text) = "" Then Exit Sub
    
    objArticulo.delete
    
    gArtMirror.deleteArticulo objArticulo, db
    
    Me.cmdEliminarArt.Enabled = False
    
End Sub

Private Sub cmdGrabar_Click()
Dim objAU As New clsDAOArticulosUbic

    If txtCodigo.Text = "" Then
        MsgBox "ERROR: Debe indicar código articulo"
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    If txtDescripcion.Text = "" Then
        MsgBox "ERROR: Debe indicar descripción articulo"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtUnitVenta.Text) Then
        MsgBox "ERROR: Debe indicar precio unitario"
        txtUnitVenta.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtUnitCompra.Text) Then txtUnitCompra.Text = 0
    
    If Me.cboProveedor.ListIndex < 0 Then
        MsgBox "ERROR: Debe Seleccionar Proveedor"
        Me.cboProveedor.SetFocus
        Exit Sub
    End If
    
    objArticulo.save
    
    gArtMirror.saveArticulo objArticulo, db
    
    objAU.artid = objArticulo.codigo
    objAU.ubicacion = Me.txtUbicacion.Text
    objAU.save
    
    gArtMirror.saveArticuloUbic objAU, db
    
    If cmdGrabar.Caption = "&Re Grabar" Then
        cmdGrabar.Caption = "&Grabar"
    Else
        objArticulo.fillListDescripcion Me.lstArticulos, Trim(Me.txtBuscar.Text)
    End If
    
    Me.txtCodigo.Text = ""
    
    Call llenarFormulario
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    
    Me.txtCodigo.Text = ""
    
    Call llenarFormulario
    
    cmdGrabar.Caption = "&Grabar"
    
End Sub

Private Sub dtpActualizacion_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.fechaactualizacion = Me.dtpActualizacion.Value
    
End Sub

Private Sub dtpUltimaCompra_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.ultimacompra = Me.dtpUltimaCompra.Value
    
End Sub

Private Sub Form_Load()
Dim strTitulos As Variant
Dim intAnchos As Variant

    strTitulos = Array("clave", "Codigo", "Descripcion", "Proveedor")
    intAnchos = Array(0, 1000, 3000, 3000)
    
    modGrid.makeGrid Me.grdAlternativo, strTitulos, intAnchos, 0, 1, flexSelectionByRow
    
    strTitulos = Array("clave", "Alias", "Proveedor", "P.Compra")
    intAnchos = Array(0, 1500, 4000, 1500)
    
    modGrid.makeGrid Me.grdAlias, strTitulos, intAnchos, 0, 1, flexSelectionByRow
    
    objProveedor.fillCombo Me.cboProveedor, db
    objProveedor.fillCombo Me.cboProveedorAlias, db
    
    Me.txtCodigo.Text = ""
    
    Me.dtpActualizacion.Value = Date
    Me.dtpUltimaCompra.Value = Date
    
    If UCase(gUsuario.nombre) = "SANJUAN" Then
        With Me
            .txtPrecioListaConIVA.Locked = True
            .txtPrecioListaSinIVA.Locked = True
            .txtPrecioVentaSinIVA.Locked = True
            .txtUnitVenta.Locked = True
            .txtUnitCompra.Locked = True
            .txtUnitCompraAnterior.Locked = True
            .txtDescripcion.Locked = True
        End With
    End If
    
    Call llenarFormulario
    
End Sub

Private Sub grdAlternativo_DblClick()
    
    If Me.grdAlternativo.row < 1 Then Exit Sub
    
    blnCargando = True
    
    objArticulo.codigo = Me.grdAlternativo.TextMatrix(Me.grdAlternativo.row, 1)
    objArticulo.findByPrimaryKey db
        
    Me.txtCodigo.Text = objArticulo.codigo
    
    Call llenarFormulario
    
    blnCargando = False

End Sub

Private Sub lstArticulos_Click()

    If Me.lstArticulos.ListIndex < 0 Then Exit Sub
    
    blnCargando = True
    
    objArticulo.findByClave Me.lstArticulos.ItemData(Me.lstArticulos.ListIndex), db
        
    Me.txtCodigo.Text = objArticulo.codigo
    
    Call llenarFormulario
    
    blnCargando = False
    
End Sub

Private Sub txtAlias_GotFocus()

    marcarseleccion Me.txtAlias
    
End Sub

Private Sub txtAlternativo_Change()

    If Trim(Me.txtAlternativo.Text) = "" Then
        Me.lstArticuloAlternativo.Clear
    Else
        objArticulo.fillListDescripcion Me.lstArticuloAlternativo, Trim(Me.txtAlternativo.Text)
    End If

End Sub

Private Sub txtAlternativo_GotFocus()

    marcarseleccion Me.txtAlternativo
    
End Sub

Private Sub txtBuscar_Change()

    If Trim(Me.txtBuscar.Text) = "" Then
        Me.lstArticulos.Clear
    Else
        objArticulo.fillListDescripcion Me.lstArticulos, Trim(Me.txtBuscar.Text)
    End If
    
End Sub

Private Sub txtBuscar_GotFocus()

    marcarseleccion Me.txtBuscar
    
End Sub

Private Sub txtCatalogo_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.catalogo = Me.txtCatalogo.Text

End Sub

Private Sub txtCatalogo_GotFocus()

    marcarseleccion Me.txtCatalogo
    
End Sub

Private Sub txtCodigo_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.codigo = Me.txtCodigo.Text
    
End Sub

Private Sub txtCodigo_GotFocus()

    marcarseleccion Me.txtCodigo
    
End Sub

Private Sub txtCodigo_KeyPress(keyAscii As Integer)

    If keyAscii = 13 Then
        keyAscii = 0
        
        Me.txtDescripcion.SetFocus
    End If
    
End Sub

Private Sub txtCodigo_LostFocus()
    
    Call llenarFormulario
    
End Sub

Private Sub txtDescripcion_Change()
    
    If blnCargando Then Exit Sub
    
    objArticulo.descripcion = Me.txtDescripcion.Text
    
End Sub

Private Sub txtDescripcion_GotFocus()

    marcarseleccion Me.txtDescripcion
    
End Sub

Private Sub txtDescuento_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.descuento = Me.txtDescuento.Text
    
End Sub

Private Sub txtDescuento_GotFocus()

    marcarseleccion Me.txtDescuento
    
End Sub

Private Sub txtMarca_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.marca = Me.txtMarca.Text
    
End Sub

Private Sub txtMarca_GotFocus()

    marcarseleccion Me.txtMarca
    
End Sub

Private Sub txtModeloCamion_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.modelocamion = Me.txtModeloCamion.Text
    
End Sub

Private Sub txtModeloCamion_GotFocus()

    marcarseleccion Me.txtModeloCamion
    
End Sub

Private Sub txtOrigen_Change()

    If blnCargando Then Exit Sub
    
    objArticulo.origen = Me.txtOrigen.Text
    
End Sub

Private Sub txtOrigen_GotFocus()

    marcarseleccion Me.txtOrigen
    
End Sub

Private Sub txtPrecioListaConIVA_Change()
Dim curAlicuotaIVA As Currency

On Error Resume Next

    If blnCargando Then Exit Sub
    
    objArticulo.preciolistaconiva = Me.txtPrecioListaConIVA.Text
    
    If vEscribiendoConIVA Then
        curAlicuotaIVA = 1.21
        If Me.chkIva105.Value = 1 Then curAlicuotaIVA = 1.105
        If Me.chkIvaExento.Value = 1 Then curAlicuotaIVA = 1
        
        Me.txtPrecioListaSinIVA.Text = Format(objArticulo.preciolistaconiva / curAlicuotaIVA, "0.00")
    
        Me.txtUtilidadLta.Text = Format((objArticulo.preciolistasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")
    End If
    
End Sub

Private Sub txtPrecioListaConIVA_GotFocus()

    marcarseleccion Me.txtPrecioListaConIVA
    
End Sub

Private Sub txtPrecioListaConIVA_KeyPress(keyAscii As Integer)

    vEscribiendoSinIVA = False
    vEscribiendoConIVA = True
    
End Sub

Private Sub txtPrecioListaSinIVA_Change()
Dim curAlicuotaIVA As Currency

On Error Resume Next

    If blnCargando Then Exit Sub
    
    objArticulo.preciolistasiniva = Me.txtPrecioListaSinIVA.Text
    
    If vEscribiendoSinIVA Then
        curAlicuotaIVA = 1.21
        If Me.chkIva105.Value = 1 Then curAlicuotaIVA = 1.105
        If Me.chkIvaExento.Value = 1 Then curAlicuotaIVA = 1
        
        Me.txtPrecioListaConIVA.Text = Format(objArticulo.preciolistasiniva * curAlicuotaIVA, "0.00")
    
        Me.txtUtilidadLta.Text = Format((objArticulo.preciolistasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")
    End If
    
End Sub

Private Sub txtPrecioListaSinIVA_GotFocus()

    marcarseleccion Me.txtPrecioListaSinIVA
    
End Sub

Private Sub txtPrecioListaSinIVA_KeyPress(keyAscii As Integer)

    vEscribiendoSinIVA = True
    vEscribiendoConIVA = False
    
End Sub

Private Sub txtPrecioVentaSinIVA_Change()
Dim curAlicuotaIVA As Currency

On Error Resume Next

    If blnCargando Then Exit Sub
    
    objArticulo.precioventasiniva = Me.txtPrecioVentaSinIVA.Text
    
    If vEscribiendoSinIVA Then
        curAlicuotaIVA = 1.21
        If Me.chkIva105.Value = 1 Then curAlicuotaIVA = 1.105
        If Me.chkIvaExento.Value = 1 Then curAlicuotaIVA = 1
        
        Me.txtUnitVenta.Text = Format(objArticulo.precioventasiniva * curAlicuotaIVA, "0.00")
    
        Me.txtUtilidadVta.Text = Format((objArticulo.precioventasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")
    End If
    
End Sub

Private Sub txtPrecioVentaSinIVA_GotFocus()

    marcarseleccion Me.txtPrecioVentaSinIVA
    
End Sub

Private Sub txtPrecioVentaSinIVA_KeyPress(keyAscii As Integer)

    vEscribiendoSinIVA = True
    vEscribiendoConIVA = False
    
End Sub

Private Sub txtProvPrecioCompra_GotFocus()

    marcarseleccion Me.txtProvPrecioCompra
    
End Sub

Private Sub txtUnitCompra_Change()

On Error Resume Next

    If blnCargando Then Exit Sub

    objArticulo.preciocomprasiniva = Val(Me.txtUnitCompra.Text)
    
    Me.txtUtilidadVta.Text = Format((objArticulo.precioventasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")
    Me.txtUtilidadLta.Text = Format((objArticulo.preciolistasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")

End Sub

Private Sub txtUnitCompra_GotFocus()

    marcarseleccion Me.txtUnitCompra
    
End Sub

Private Sub txtUnitCompraAnterior_Change()

On Error Resume Next

    If blnCargando Then Exit Sub

    objArticulo.preciocomprasinivaanterior = Val(Me.txtUnitCompraAnterior.Text)

End Sub

Private Sub txtUnitCompraAnterior_GotFocus()

    marcarseleccion Me.txtUnitCompraAnterior
    
End Sub

Private Sub txtUnitVenta_Change()
Dim curAlicuotaIVA As Currency

On Error Resume Next

    If blnCargando Then Exit Sub

    objArticulo.precioventaconiva = Val(Me.txtUnitVenta.Text)
    
    If vEscribiendoConIVA Then
        curAlicuotaIVA = 1.21
        If Me.chkIva105.Value = 1 Then curAlicuotaIVA = 1.105
        If Me.chkIvaExento.Value = 1 Then curAlicuotaIVA = 1
        
        Me.txtPrecioVentaSinIVA.Text = Format(objArticulo.precioventaconiva / curAlicuotaIVA, "0.00")
        
        Me.txtUtilidadVta.Text = Format((objArticulo.precioventasiniva / objArticulo.preciocomprasiniva - 1) * 100, "0.00")
    End If
    
End Sub

Private Sub txtUnitVenta_GotFocus()

    marcarseleccion Me.txtUnitVenta
    
End Sub

Private Sub llenarFormulario()
Dim objAAP As New clsVDAOAlterArtProv

Dim objAPr As New clsVDAOAliasProv

Dim objDetArtic As New clsDAODetartic

Dim objAU As New clsDAOArticulosUbic

On Error Resume Next

    Me.cmdEliminarArt.Enabled = False

    blnCargando = True

    Me.cboProveedor.ListIndex = -1

    With objArticulo
        .codigo = Trim(Me.txtCodigo.Text)
        .findByPrimaryKey db
        
        Me.txtMovs.Text = objDetArtic.cantidadMovs(.codigo, db)
        
        Me.cmdEliminarArt.Enabled = IIf(Val(Me.txtMovs.Text) > 0, False, True)
        
        If .clave = 0 Then
            objArticulosAlias.findByAlias Me.txtCodigo.Text, db
            
            If objArticulosAlias.clave <> 0 Then
                .codigo = objArticulosAlias.artid
                .findByPrimaryKey db
            End If
        End If
    
        Me.txtCodigo.Text = .codigo
        Me.txtDescripcion.Text = .descripcion
        Me.txtUnitVenta.Text = Format(.precioventaconiva, "0.00")
        Me.txtPrecioVentaSinIVA.Text = Format(.precioventasiniva, "0.00")
        Me.txtPrecioListaConIVA.Text = Format(.preciolistaconiva, "0.00")
        Me.txtPrecioListaSinIVA.Text = Format(.preciolistasiniva, "0.00")
        
        Me.txtUtilidadVta.Text = Format((.precioventasiniva / .preciocomprasiniva - 1) * 100, "0.00")
        Me.txtUtilidadLta.Text = Format((.preciolistasiniva / .preciocomprasiniva - 1) * 100, "0.00")
        
        Me.txtUnitCompra.Text = Format(.preciocomprasiniva, "0.00")
        vPrecioCompraOriginal = .preciocomprasiniva
        Me.txtUnitCompraAnterior.Text = Format(.preciocomprasinivaanterior, "0.00")
        Me.txtModeloCamion.Text = .modelocamion
        Me.txtDescuento.Text = .descuento
        Me.txtOrigen.Text = .origen
        Me.chkIvaExento.Value = .exento
        Me.chkIva105.Value = .iva105
        
        objProveedor.proveedorID = .prvID
        objProveedor.findByPrimaryKey db
        Me.cboProveedor.Text = objProveedor.comboText
                
        Me.dtpActualizacion.Value = .fechaactualizacion
        Me.dtpUltimaCompra.Value = .ultimacompra
        
        objAU.artid = .codigo
        objAU.findByPrimaryKey
        Me.txtUbicacion.Text = objAU.ubicacion
        
        Me.txtMarca.Text = .marca
        Me.txtCatalogo.Text = .catalogo
        
        objAAP.llenarGrilla Me.grdAlternativo, .codigo, db
        objAPr.llenarGrilla Me.grdAlias, .codigo, db
        
        If .exist Then cmdGrabar.Caption = "&Re Grabar"
        
    End With
    
    blnCargando = False
    
End Sub

Private Sub txtUnitVenta_KeyPress(keyAscii As Integer)

    vEscribiendoSinIVA = False
    vEscribiendoConIVA = True
    
End Sub
