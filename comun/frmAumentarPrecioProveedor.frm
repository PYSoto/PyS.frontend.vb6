VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAumentarPrecioProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Precios por Proveedor"
   ClientHeight    =   6390
   ClientLeft      =   2505
   ClientTop       =   1860
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7950
   Begin VB.TextBox txtProveedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   79429633
      CurrentDate     =   41005
   End
   Begin VB.CheckBox chkAlias 
      Caption         =   "incluir Alias"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   288
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdArticulos 
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   6000
      TabIndex        =   7
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Cambiar"
      Height          =   372
      Left            =   4080
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdAlias 
      Height          =   2055
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Actualización"
      Height          =   195
      Left            =   4080
      TabIndex        =   11
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Porcentaje Variación"
      Height          =   195
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Artículos"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAumentarPrecioProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private proveedor As New clsDAOProveedor
Private articulo As New clsDAOArticulos

Private articuloalias As New clsDAOArticulosAlias

Private articulos As New Collection
Private articulosalias As New Collection

Private Sub cleanform()
    
    Me.grdArticulos.Rows = 1
    Me.grdAlias.Rows = 1

    Set proveedor = New clsDAOProveedor
    If proveedorglobal.proveedorID > 0 Then Set proveedor = proveedorglobal
    Me.txtProveedor.Text = proveedor.textSearch
    
    fillArticulos
    
    If Me.chkAlias.Value = 1 Then fillArticulosAlias

End Sub

Private Sub fillArticulos()

    If proveedor.proveedorID = 0 Then Exit Sub

    Set articulos = articulo.collectionByProveedor(proveedor.proveedorID)
    
    Me.grdArticulos.Rows = 1
    Me.grdArticulos.Redraw = False
    For Each articulo In articulos
        Me.grdArticulos.AddItem modGrid.array2itemGrid(Array(articulo.codigo, articulo.descripcion, Format(articulo.preciocomprasiniva, "0.00"), Format(articulo.precioventasiniva, "0.00"), Format(articulo.preciolistasiniva, "0.00")))
    Next
    Me.grdArticulos.Redraw = True
    
End Sub

Private Sub fillArticulosAlias()
Dim articulolocal As New clsDAOArticulos

    If proveedor.proveedorID = 0 Then Exit Sub

    Set articulosalias = articuloalias.collectionByProveedor(proveedor.proveedorID)
    
    Me.grdAlias.Rows = 1
    Me.grdAlias.Redraw = False
    For Each articuloalias In articulosalias
        articulolocal.codigo = articuloalias.artid
        articulolocal.findByPrimaryKey db
        
        Me.grdAlias.AddItem modGrid.array2itemGrid(Array(articuloalias.artid, articuloalias.alias, articulolocal.descripcion, Format(articuloalias.preciocompra, "0.00")))
    Next
    Me.grdAlias.Redraw = True
    
End Sub

Private Sub chkAlias_Click()

    Me.grdAlias.Rows = 1
    
    If proveedor.proveedorID = 0 Then Exit Sub
    
    If Me.chkAlias.Value = 1 Then fillArticulosAlias
    
End Sub

Private Sub cmdCambiar_Click()

    If Val(Me.txtPorcentaje.Text) = 0 Then Exit Sub
    
    Me.MousePointer = 11
    
    Me.cmdCambiar.Enabled = False
    
    For Each articulo In articulos
        articulo.aumentarPrecios Val(Me.txtPorcentaje.Text), Me.dtpFecha.Value, db
        gArtMirror.saveArticulo articulo, db
    Next
    
    If Me.chkAlias.Value = 1 Then
        For Each articuloalias In articulosalias
            articuloalias.aumentarPrecios Val(Me.txtPorcentaje.Text), db
            gArtMirror.saveArticuloAlias articuloalias, db
        Next
    End If
    
    Me.cmdCambiar.Enabled = True
    
    Me.MousePointer = 0
    
    cleanform
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.dtpFecha.Value = Date

    modGrid.makeGrid2 Me.grdArticulos, Array(Array("Codigo", 1000), Array("Descripcion", 3000), Array("Compra", 1000), Array("Venta", 1000), Array("Lista", 1000)), 0, 1, flexSelectionByRow
    modGrid.makeGrid2 Me.grdAlias, Array(Array("Codigo", 1000), Array("Alias", 1000), Array("Descripcion", 4000), Array("Compra", 1000)), 0, 1, flexSelectionByRow
    
    cleanform
    
End Sub

Private Sub txtPorcentaje_GotFocus()

    marcarseleccion Me.txtPorcentaje
    
End Sub

Private Sub txtPorcentaje_LostFocus()

    If Not IsNumeric(Me.txtPorcentaje.Text) Then Me.txtPorcentaje.Text = ""
    
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
    
    cleanform
    
End Sub

