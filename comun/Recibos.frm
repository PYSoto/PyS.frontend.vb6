VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Recibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   6270
   ClientLeft      =   1530
   ClientTop       =   1035
   ClientWidth     =   9840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9840
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   3000
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtTitular 
      Height          =   285
      Left            =   2160
      TabIndex        =   24
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   6000
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtImporte 
      Height          =   285
      Left            =   4080
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox cboValor 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox txtTotalValores 
      Height          =   288
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtPago 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdComprobantes 
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.TextBox txtComprobante 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtPrefijo 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cboComprobante 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtReserva 
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      TabIndex        =   9
      Top             =   5760
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
      Left            =   7920
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datFecha 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100794369
      CurrentDate     =   39075
   End
   Begin MSComCtl2.DTPicker datEmision 
      Height          =   375
      Left            =   7920
      TabIndex        =   26
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100794369
      CurrentDate     =   39074
   End
   Begin MSFlexGridLib.MSFlexGrid grdValores 
      Height          =   1095
      Left            =   240
      TabIndex        =   27
      Top             =   4560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1931
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker datVto 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100794369
      CurrentDate     =   39074
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   9240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblBanco 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      Height          =   195
      Left            =   3000
      TabIndex        =   36
      Top             =   3960
      Width           =   465
   End
   Begin VB.Label lblTitular 
      AutoSize        =   -1  'True
      Caption         =   "Titular"
      Height          =   195
      Left            =   2160
      TabIndex        =   35
      Top             =   3960
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   3960
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Emisión"
      Height          =   195
      Left            =   7920
      TabIndex        =   33
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   6000
      TabIndex        =   32
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
      Height          =   195
      Left            =   4080
      TabIndex        =   31
      Top             =   3360
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   7920
      TabIndex        =   29
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comprobante"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total"
      Height          =   195
      Left            =   4080
      TabIndex        =   13
      Top             =   720
      Width           =   930
   End
End
Attribute VB_Name = "Recibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objCliente As New clsDAOClientes

Private objTiposComprob As New clsDAOTiposComprob

Private objCgoValor As New clsDAOCgosValores

Private objMovclie As New clsDAOMovclie

Private vTipocompro As String

Private vDebita As Boolean

Private vRow As Integer

Private claveMovCli As Long

Private Sub buscaFacPen()
Dim intRow As Integer

Dim dblSaldo As Double

Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    Me.grdComprobantes.Rows = 1

    Me.MousePointer = 11
    
    strSQL = "SELECT movclie.*, tiposcomprob.* FROM movclie, tiposcomprob"
    strSQL = strSQL & " WHERE cgoclie = " & objCliente.codigo
    strSQL = strSQL & " AND tiposcomprob.codigo = movclie.cgocomprob"
    strSQL = strSQL & ";"
    
    Set rstQuery = db.query(strSQL)
    
    Do While Not rstQuery.EOF
    
        dblSaldo = rstQuery!importe - rstQuery!cancelado
        
        If dblSaldo <> 0 And rstQuery!aplicable <> 0 Then
        
            With Me.grdComprobantes
                .Rows = .Rows + 1
                
                intRow = .Rows - 1
                
                .TextMatrix(intRow, 0) = rstQuery!clave
                .TextMatrix(intRow, 1) = rstQuery!descripcion
                .TextMatrix(intRow, 2) = rstQuery!prefijo
                .TextMatrix(intRow, 3) = rstQuery!nroComprob
                .TextMatrix(intRow, 4) = rstQuery!fechacomprob
                .TextMatrix(intRow, 5) = Format(rstQuery!importe, "0.00")
                .TextMatrix(intRow, 6) = Format(rstQuery!cancelado, "0.00")
                .TextMatrix(intRow, 7) = Format(dblSaldo, "0.00")
                .TextMatrix(intRow, 8) = Format(0, "0.00")
                
            End With
            
        End If
        
        rstQuery.MoveNext
        
    Loop
    
    rstQuery.Close
    
    Me.MousePointer = 0

End Sub

Public Sub grabaFacPen()
Dim dblPago As Currency

Dim strSQL As String

Dim intRow As Integer

Dim objMovCli As New clsDAOMovclie

Dim objFactCobradas As New clsDAOFactcobradas

    With Me.grdComprobantes
    
        For intRow = 1 To .Rows - 1
        
            dblPago = Val(Me.grdComprobantes.TextMatrix(intRow, 8))
            
            If dblPago <> 0 Then
            
                objMovCli.clave = Val(Me.grdComprobantes.TextMatrix(intRow, 0))
                objMovCli.findByPrimaryKey db
                objMovCli.cancelado = objMovCli.cancelado + dblPago
                objMovCli.save db
                
                objFactCobradas.clavemovc = Val(.TextMatrix(intRow, 0))
                objFactCobradas.clavepago = claveMovCli
                objFactCobradas.importe = dblPago
                objFactCobradas.add
            
            End If
            
        Next intRow
        
    End With
    
End Sub

Public Sub grabaMovCli()
Dim intFactor As Integer
Dim intPrefijo As Integer

Dim strSQL As String
Dim rstQuery As ADODB.Recordset

    intFactor = IIf(vDebita, 1, -1)
    
    intPrefijo = 0
    
    If cmdIngresar.Caption = "Ingresar" Then intPrefijo = txtPrefijo.Text
    
    Set objMovclie = New clsDAOMovclie

    With objMovclie
        .cgoclie = objCliente.codigo
        .cgocomprob = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
        .fechacomprob = Me.datFecha.value
        .importe = Val(txtTotal.Text) * intFactor
        .prefijo = intPrefijo
        .nroComprob = Val(Me.txtComprobante.Text)
        .recibo = 1
        
        .add db
        
        claveMovCli = .clave
        
    End With

    Me.txtClave.Text = claveMovCli
    
End Sub

Public Sub borrar()

    txtPrefijo.Text = ""
    txtComprobante.Text = ""
    txtTotal.Text = ""
    txtReserva.Text = ""
    Me.datFecha.value = Date
    Me.datEmision.value = Date
    Me.datVto.value = Date
    vTipocompro = " "
    cmdIngresar.Enabled = True
    
    Me.grdValores.Rows = 1
    
    Me.txtTotalValores.Text = ""
    Me.txtImporte.Text = ""
    Me.txtNumero.Text = ""
    Me.txtBanco.Text = ""
    Me.txtTitular.Text = ""
    
    Me.txtPago.Visible = False
    
End Sub

Private Sub configura()

    If Me.cboComprobante.ListIndex < 0 Then Exit Sub
    
    vDebita = False

    objTiposComprob.codigo = Me.cboComprobante.ItemData(Me.cboComprobante.ListIndex)
    objTiposComprob.findByPrimaryKey db
    
    If objTiposComprob.exist(db) Then
        vDebita = IIf(objTiposComprob.debita <> 0, True, False)
        
        Me.txtPrefijo.Text = objTiposComprob.puntovta

        If objTiposComprob.aplicapend <> 0 Then
            Me.grdComprobantes.Enabled = True
            txtTotal.Enabled = False
        Else
            Me.grdComprobantes.Enabled = False
            txtTotal.Enabled = True
        End If
        
        If objTiposComprob.ctacte <> 0 Then
            Label7.Visible = True
            txtPrefijo.Visible = True
            cmdIngresar.Caption = "Ingresar"
        Else
            Label7.Visible = False
            txtPrefijo.Visible = False
            cmdIngresar.Caption = "Aplicar"
        End If
    End If
    
End Sub

Private Sub cboComprobante_Click()

    If cboComprobante.ListIndex < 0 Then Exit Sub
        
    Call configura
    
    With objTiposComprob
        Me.txtComprobante.Text = objMovclie.buscaNumeroRecibo(.puntovta, .tipocomprob, .codigo, db)
    End With
    
    Me.cmdIngresar.Enabled = True
    
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
Dim impresion_service As New clsCtlImpresion
    
    If objCliente.codigo = 0 Then Exit Sub
    
    If Me.cboComprobante.ListIndex < 0 Then Exit Sub
    
    If txtPrefijo.Visible And Not IsNumeric(txtPrefijo.Text) Then
        MsgBox "Debe indicar Número del Prefijo"
        txtPrefijo.SetFocus
        Exit Sub
    End If
    
    If txtComprobante.Visible And Not IsNumeric(txtComprobante.Text) Then
        MsgBox "Debe indicar Número del Comprobante"
        txtComprobante.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtTotal.Text) Then
        MsgBox "Debe Indicar el Importe Total"
        Exit Sub
    Else
        If txtTotal.Text <> 0 And cmdIngresar.Caption = "Aplicar" Then
            MsgBox "El Importe Total debe ser cero"
            Exit Sub
        End If
        If txtTotal.Text = 0 And cmdIngresar.Caption = "Ingresar" Then
            MsgBox "El Importe Total debe ser distinto de cero"
            Exit Sub
        End If
    End If
    
    If Not IsNumeric(txtReserva.Text) Then txtReserva.Text = 0
    
    If Val(txtTotal.Text) < Val(txtReserva.Text) Then
        MsgBox "El Importe Total debe ser igual o mayor al minimo de la reserva para bloquear definitivamente"
        Exit Sub
    End If
    
    If cmdIngresar.Caption = "Ingresar" Then
        If Val(Me.txtTotalValores.Text) > 0 And Val(Me.txtTotalValores.Text) <> Val(Me.txtTotal.Text) Then
            MsgBox "ERROR: No coinciden Totales"
            Exit Sub
        End If
        
        If Me.grdValores.Rows = 1 Then
            MsgBox "ERROR: Faltan VALORES"
            Exit Sub
        End If
        
        Call grabaMovCli
        Call grabaFacPen
        Call grabaValores
    Else
        Call grabaMovCli
        Call grabaFacPen
    End If
    
    objMovclie.clave = claveMovCli
    objMovclie.findByPrimaryKey db
    
    If objMovclie.recibo <> 0 Then impresion_service.printReport Me.crpConsulta, "rptRecibo", db.sconection, Array("FactPagadas", "Valores"), Array(Array("pClave", objMovclie.clave))
    
    Call cmdCancelar_Click
    
End Sub

Private Sub cmdCancelar_Click()

    Call borrar
    
    Call buscaFacPen
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim strTitulos As Variant
Dim intAnchos As Variant

    strTitulos = Array("Codigo", "Concepto", "Numero", "Importe", "Titular", "Banco", "Emision", "Vencimiento")
    
    intAnchos = Array(0, 3000, 1000, 1000, 1000, 1000, 1000, 1000)
    
    modGrid.makeGrid Me.grdValores, strTitulos, intAnchos, 0, 1, flexSelectionByRow
    
    strTitulos = Array("Clave", "Comprob", "Pref", "Nro.Comp", "Fecha", "Total", "A Cta", "Saldo", "Pago")
    
    intAnchos = Array(0, 2000, 500, 1000, 1000, 1000, 1000, 1000, 1000)
    
    modGrid.makeGrid Me.grdComprobantes, strTitulos, intAnchos, 0, 1, flexSelectionFree

    Me.datFecha.value = Date
    
    Call borrar
    
    objCgoValor.fillCombo Me.cboValor
    
End Sub

Private Sub grdComprobantes_Click()
Dim intRow As Integer
Dim intCol As Integer

    intRow = Me.grdComprobantes.row
    intCol = Me.grdComprobantes.Col

    If intRow = 0 Then Exit Sub
    
    If intCol <> 8 Then Exit Sub
    
    vRow = intRow
    
    modGrid.setTextBox Me.grdComprobantes, Me.txtPago
    
    With Me.txtPago
        .Text = Format(Val(Me.grdComprobantes.TextMatrix(vRow, 8)), "0.00")
        
        .Visible = True
    End With
    
End Sub

Private Sub grdComprobantes_Scroll()

    Me.txtPago.Visible = False
    
End Sub

Private Sub txtCliente_GotFocus()

    marcarseleccion Me.txtCliente
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)

    objCliente.formBuscar frmBuscar, objCliente, KeyAscii, "Clientes"
    
    Me.txtCliente.Text = objCliente.descripcionBuscar
    KeyAscii = 0

    If objCliente.posicion = 1 Or objCliente.posicion = 4 Then
         vTipocompro = "A"
    Else
         vTipocompro = "B"
    End If
    
    objTiposComprob.fillComboRecibos Me.cboComprobante, vTipocompro, db
    
    Call buscaFacPen

End Sub

Private Sub txtComprobante_GotFocus()

    marcarseleccion Me.txtComprobante
    
End Sub

Private Sub txtPago_GotFocus()

    marcarseleccion Me.txtPago
    
End Sub

Private Sub txtPago_LostFocus()

    Me.grdComprobantes.TextMatrix(vRow, 8) = Format(Val(Me.txtPago.Text), "0.00")
    
    Me.txtPago.Visible = False
    
    Call calcularTotal
    
End Sub

Private Sub calcularTotal()
Dim dblTotal As Double
Dim dblPago As Double
Dim dblSaldo As Double

Dim intRow As Integer

    dblTotal = 0
    
    With Me.grdComprobantes
        
        For intRow = 1 To .Rows - 1
            
            dblPago = Val(.TextMatrix(intRow, 8))
            dblSaldo = Val(.TextMatrix(intRow, 7))
            
            dblTotal = dblTotal + dblPago
            
            If dblPago > dblSaldo And dblSaldo > 0 Then MsgBox "Excede Monto dblSaldo"
            If dblPago < 0 And dblSaldo > 0 Then MsgBox "Debe ingresar monto positivo"
            If dblPago < dblSaldo And dblSaldo < 0 Then MsgBox "Importe Menor al Crédito"
            If dblPago > 0 And dblSaldo < 0 Then MsgBox "Debe ingresar monto negativo"
            
        Next intRow
        
    End With
    
    txtTotal.Text = Format(dblTotal, "0.00")
   
End Sub

Private Function calcularValores() As Double
Dim dblTotal As Double

Dim intCiclo As Integer

    dblTotal = 0
    
    For intCiclo = 1 To Me.grdValores.Rows - 1
        dblTotal = dblTotal + Val(Me.grdValores.TextMatrix(intCiclo, 3))
    Next intCiclo
    
    calcularValores = dblTotal
    
End Function

Private Sub grabaValores()
Dim intCiclo As Integer

Dim objValores As clsDAOValores

    For intCiclo = 1 To Me.grdValores.Rows - 1
        Set objValores = New clsDAOValores
        
        With objValores
            .codigo = Me.grdValores.TextMatrix(intCiclo, 0)
            .cgocli = objCliente.codigo
            .fechaemi = CDate(Me.grdValores.TextMatrix(intCiclo, 6))
            .fechavto = CDate(Me.grdValores.TextMatrix(intCiclo, 7))
            .nroComprob = Val(Me.grdValores.TextMatrix(intCiclo, 2))
            .importe = Val(Me.grdValores.TextMatrix(intCiclo, 3))
            
            objCgoValor.codigo = .codigo
            objCgoValor.findByPrimaryKey
            
            .clavemovv = claveMovCli
            .titular = Me.grdValores.TextMatrix(intCiclo, 4)
            .banco = Me.grdValores.TextMatrix(intCiclo, 5)
            
            .add
        End With
        
    Next intCiclo
    
End Sub

