VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPresupuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   6015
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11790
   Begin VB.CommandButton cmdActualizarPrecios 
      Caption         =   "Actualizar Precios"
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox chkCtaCte 
      Caption         =   "Cuenta Corriente"
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CheckBox chkRemito 
      Caption         =   "Remito"
      Height          =   375
      Left            =   4080
      TabIndex        =   37
      Top             =   5520
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpPresupuesto 
      Left            =   9720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6000
      TabIndex        =   36
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdPendientes 
      Caption         =   "Pendientes"
      Height          =   255
      Left            =   9840
      TabIndex        =   35
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   34
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   33
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txtIVA 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5040
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
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtNro 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   5535
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
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtUnitario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtArtID 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   255
      Left            =   9840
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   255
      Left            =   9840
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   270
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   735
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpActualizacion 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   80478209
      CurrentDate     =   40182
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetalle 
      Height          =   2055
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3625
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker datFecha 
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   600
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
      Format          =   80478209
      CurrentDate     =   39075
   End
   Begin MSComCtl2.DTPicker datFechaVto 
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   600
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
      Format          =   80478209
      CurrentDate     =   40909
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "I.V.A."
      Height          =   195
      Left            =   7920
      TabIndex        =   31
      Top             =   4800
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Neto"
      Height          =   195
      Left            =   6000
      TabIndex        =   30
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   9840
      TabIndex        =   29
      Top             =   4800
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   " Artículos "
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Left            =   4080
      TabIndex        =   22
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   6960
      TabIndex        =   20
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ingreso"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   525
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "P.U.s/IVA"
      Height          =   195
      Left            =   7920
      TabIndex        =   18
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   7920
      TabIndex        =   17
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   840
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   0
      Top             =   1320
      Width           =   11775
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
      Height          =   195
      Left            =   3600
      TabIndex        =   15
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Ultima Act"
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Observac"
      Height          =   195
      Left            =   6960
      TabIndex        =   13
      Top             =   480
      Width           =   690
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objAr As New clsDAOArticulos

Private objPr As New clsDAOPresupuesto

Private colDP As Collection
Private colDPx As Collection

Private vOrden As Integer
Private vSelect As Integer

Private vDescuento As Currency

Private Sub clean()

    Set objPr = New clsDAOPresupuesto
    
    objPr.newID True
    
    Set colDP = New Collection
    
    vOrden = 1
    vSelect = 0

    Me.cboCliente.ListIndex = -1
    Me.grdDetalle.Rows = 1
    txtTotal.Text = ""
    txtNeto.Text = ""
    txtIVA.Text = ""
    txtCantidad.Text = ""
    txtUnitario.Text = ""
    Me.chkCtaCte.Value = 0
    Me.txtObservaciones.Text = ""
    
    Me.datFecha.Value = Date
    Me.datFechaVto.Value = Date
    cmdIngresar.Caption = "Ingresar"
    cmdAnular.Visible = False
    cmdGrabar.Caption = "Grabar"
    
    Me.txtNro.Text = "Nuevo"
    
End Sub

Private Sub fillArticulo()
Dim objAU As New clsDAOArticulosUbic
     
    Me.dtpActualizacion.Value = Date
    
    With objAr
        Me.dtpActualizacion.Value = .fechaactualizacion
        Me.txtUnitario.Text = Format(.precioventasiniva * vDescuento, "0.00")
    End With
    With objAU
        .artID = objAr.codigo
        .findByPrimaryKey
        Me.txtUbicacion.Text = .ubicacion
    End With

End Sub

Private Sub fillGrid()
Dim objDP As clsDAODetpresupuesto

Dim curNeto As Currency
Dim curTotal As Currency

    curNeto = 0
    curTotal = 0
    vOrden = 1

    Set colDPx = New Collection
    
    Me.grdDetalle.Rows = 1
    Me.grdDetalle.Redraw = False
    For Each objDP In colDP
        With objDP
            objAr.codigo = .artID
            objAr.findByPrimaryKey db
            Me.grdDetalle.AddItem modGrid.array2itemGrid(Array(.artID, objAr.descripcion, Format(.unitsiva, "0.00"), .cantidad, Format(.unitsiva * .cantidad, "0.00")))
            Me.grdDetalle.RowData(Me.grdDetalle.Rows - 1) = .orden
            curNeto = curNeto + .cantidad * .unitsiva
            curTotal = curTotal + .cantidad * .unitciva
            vOrden = vOrden + 1
            
            .orden = vOrden - 1
        End With
        colDPx.add objDP, "k." & (vOrden - 1)
    Next
    Me.grdDetalle.Redraw = True
    
    Set colDP = colDPx
    
    Me.txtNeto.Text = Format(curNeto, "0.00")
    Me.txtTotal.Text = Format(curTotal, "0.00")
    Me.txtIVA.Text = Format(curTotal - curNeto, "0.00")
    
End Sub

Private Sub cboArticulo_Click()

    If Me.cboArticulo.ListIndex < 0 Then Exit Sub
    
    With objAr
        .findByClave Me.cboArticulo.ItemData(Me.cboArticulo.ListIndex), db
        
        Me.txtArtID.Text = .codigo
        
        fillArticulo
    End With
    
End Sub

Private Sub cboArticulo_GotFocus()

    SendMessageLong Me.cboArticulo.hwnd, &H14F, True, 0

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
        
        fillArticulo
    End With

End Sub

Private Sub cboCliente_Click()
Dim objCli As New clsDAOClientes

    vDescuento = 0

    If cboCliente.ListIndex < 0 Then Exit Sub
    
    objCli.findByCodigo Me.cboCliente.ItemData(Me.cboCliente.ListIndex), db
    
    vDescuento = 1 - objCli.descuento / 100
    
End Sub

Private Sub cboCliente_GotFocus()
    
    SendMessageLong Me.cboCliente.hwnd, &H14F, True, 0
    
End Sub

Private Sub cmdActualizarPrecios_Click()
Dim objDP As clsDAODetpresupuesto

Dim objAr As New clsDAOArticulos

    For Each objDP In colDP
        With objAr
            .codigo = objDP.artID
            .findByPrimaryKey db
        End With
        With objDP
            .unitsiva = objAr.precioventasiniva * vDescuento
            .unitciva = objAr.precioventaconiva * vDescuento
        End With
    Next
    
    fillGrid
    
End Sub

Private Sub cmdImprimir_Click()
Dim ctlImp As New clsCtlImpresion
        
    Me.cmdImprimir.Enabled = False
    
    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpPresupuesto, "rptPresupuesto", db.sconection, , Array(Array("pPreID", objPr.preID), Array("pRemito", Me.chkRemito.Value))

    Me.MousePointer = 0

    Me.cmdImprimir.Enabled = True

End Sub

Private Sub cmdIngresar_Click()
Dim objDP As New clsDAODetpresupuesto
    
    If Me.cboCliente.ListIndex < 0 Then
        MsgBox "ERROR: Falta CLIENTE"
        Exit Sub
    End If
    
    If colDP.Count = 0 Then
        MsgBox "ERROR: Sin ARTICULOS"
        Exit Sub
    End If
    
    With objPr
        If Me.txtNro.Text = "Nuevo" Then .newID True
    
        .cliID = Me.cboCliente.ItemData(Me.cboCliente.ListIndex)
        .fecha = Me.datFecha.Value
        .fechavto = Me.datFechaVto.Value
        .observac = Me.txtObservaciones.Text
        .ctacte = Me.chkCtaCte.Value
        
        .save
    End With
    
    vOrden = 1
    For Each objDP In colDP
        With objDP
            .preID = objPr.preID
            .orden = vOrden
            
            .save
        End With
        vOrden = vOrden + 1
    Next
    
    objDP.deleteRest objPr.preID, vOrden, db
    
    cmdImprimir_Click
    
    cmdCancelar_Click
    
End Sub

Private Sub cmdCancelar_Click()

    clean
    
    cmdIngresar.Enabled = True
    
End Sub

Private Sub cmdPendientes_Click()
Dim blnCargar As Boolean

Dim objCl As New clsDAOClientes

Dim objDP As New clsDAODetpresupuesto

    blnCargar = False
    
    Load frmPresupPendiente
    
    frmPresupPendiente.Show vbModal
    
    If Not IsNull(frmPresupPendiente.preID) Then
        objPr.preID = frmPresupPendiente.preID
        objPr.findByPrimaryKey
        
        objCl.findByCodigo objPr.cliID, db
        If objCl.facturable = 0 Then
            MsgBox "ERROR: Cliente NO FACTURABLE"
            blnCargar = False
        Else
            blnCargar = True
            If objCl.negID = 0 Then
                objCl.negID = gEmpresa.negID
                objCl.save db
            End If
        End If
    End If
    
    Unload frmPresupPendiente
    
    If blnCargar Then
        Me.txtNro.Text = objPr.preID
        Me.cboCliente.Text = objCl.comboText
        Me.datFecha.Value = objPr.fecha
        Me.datFechaVto.Value = objPr.fechavto
        Me.txtObservaciones.Text = objPr.observac
        Me.chkCtaCte.Value = objPr.ctacte
        
        Set colDP = objDP.collectionByPreID(objPr.preID)
        
        fillGrid
    End If
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cmdGrabar_Click()
Dim objDP As New clsDAODetpresupuesto

    If Me.cboCliente.ListIndex < 0 Then
        MsgBox "ERROR: Debe Indicar CLIENTE"
        Me.cboCliente.SetFocus
        Exit Sub
    End If
    
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
    
    objAr.codigo = Me.txtArtID.Text
    objAr.findByPrimaryKey db
    
    If objAr.clave = 0 Then
        MsgBox "ERROR: Artículo NO REGISTRADO"
        Me.txtArtID.SetFocus
        Exit Sub
    End If
    
    If vSelect = 0 Then
        vSelect = vOrden
    Else
        Set objDP = colDP("k." & vSelect)
    End If
    
    With objDP
        .orden = vSelect
        .artID = objAr.codigo
        .cantidad = Val(Me.txtCantidad.Text)
        .unitsiva = objAr.precioventasiniva * vDescuento
        .unitciva = objAr.precioventaconiva * vDescuento
    End With
    
    If vSelect = vOrden Then colDP.add objDP, "k." & vSelect
    
    fillGrid
    
    txtArtID.SetFocus
    
    vSelect = 0
    
    Me.cmdGrabar.Caption = "Grabar"
    
End Sub

Private Sub cmdAnular_Click()

    vSelect = 0
    
    If Me.grdDetalle.Row < 1 Then Exit Sub
    
    colDP.Remove "k." & Me.grdDetalle.RowData(Me.grdDetalle.Row)
    
    fillGrid
    
    cmdAnular.Visible = False
    cmdGrabar.Caption = "Grabar"
    
End Sub

Private Sub Form_Load()
Dim objCl As New clsDAOClientes
    
    Me.datFecha.Value = Date
    Me.datFechaVto.Value = Date
    
    vDescuento = 0

    modGrid.makeGrid2 Me.grdDetalle, Array(Array("Codigo", 1500), Array("Descripcion", 5000), Array("P.Unitario", 1400), Array("Cantidad", 1400), Array("Total", 1400)), 0, 1, flexSelectionByRow
    
    objCl.fillComboFacturable Me.cboCliente, db
    
    clean
    
End Sub

Private Sub grdDetalle_DblClick()
Dim objDP As clsDAODetpresupuesto

    If Me.grdDetalle.Row < 1 Then Exit Sub
    
    vSelect = Me.grdDetalle.RowData(Me.grdDetalle.Row)
    
    Set objDP = colDP("k." & vSelect)
    
    objAr.codigo = objDP.artID
    objAr.findByPrimaryKey db
    Me.txtArtID.Text = objAr.codigo
    Me.cboArticulo.Text = objAr.comboText
    
    fillArticulo
    
    Me.txtCantidad.Text = objDP.cantidad
    Me.txtUnitario.Text = Format(objDP.unitsiva, "0.00")

    Me.cmdGrabar.Caption = "Regrabar"
    Me.cmdAnular.Visible = True
    
End Sub


Private Sub txtArtID_GotFocus()

    marcarseleccion Me.txtArtID
    
End Sub

Private Sub txtObservaciones_GotFocus()

    marcarseleccion Me.txtObservaciones
    
End Sub

Private Sub txtCantidad_GotFocus()

    marcarseleccion Me.txtCantidad

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdGrabar_Click
    End If
    
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
Dim objAA As New clsDAOArticulosAlias

    If txtArtID.Text = "" Then Exit Sub
    
    With objAr
        .codigo = Me.txtArtID.Text
        .findByPrimaryKey db
        
        If .clave = 0 Then
            objAA.findByAlias Me.txtArtID.Text, db
            
            If objAA.clave <> 0 Then
                .codigo = objAA.artID
                .findByPrimaryKey db
            End If
            Me.txtArtID.Text = .codigo
        End If
        
        If .descripcion = "" Then
            Me.txtArtID.Text = ""
            Me.cboArticulo.ListIndex = -1
            Exit Sub
        End If
        
        If Me.cboArticulo.Text <> .comboText Then cboArticulo.Text = .comboText
            
        fillArticulo
    End With
    
End Sub

