VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form LCtaCtecli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Corrientes Clientes"
   ClientHeight    =   2895
   ClientLeft      =   2430
   ClientTop       =   1950
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdLista 
      Caption         =   "Comprobantes"
      Height          =   372
      Left            =   3120
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chkSinSaldo 
      Caption         =   "Sin Saldo Anterior"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin VB.TextBox txtDesde 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin Crystal.CrystalReport crpClientes 
      Left            =   240
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cboClienteD 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta Fecha :"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde Fecha :"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "LCtaCtecli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objCliente As New clsDAOClientes

Private objMovclie As New clsDAOMovclie

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cboClienteD_Click()

    If cboClienteD.ListIndex < 0 Then Exit Sub
    
    Me.txtDesde.Text = cboClienteD.ItemData(cboClienteD.ListIndex)
    
End Sub

Private Sub cmdConsultar_Click()
Dim curSaldoDeudor As Currency

Dim impresion_service As New clsCtlImpresion

    If Me.cboClienteD.ListIndex < 0 Then Exit Sub

    Me.cmdConsultar.Enabled = False
    
    Me.MousePointer = 11

    curSaldoDeudor = 0
    
    If Me.chkSinSaldo.value = 0 Then curSaldoDeudor = objMovclie.saldoImporteHasta(Me.cboClienteD.ItemData(Me.cboClienteD.ListIndex), Me.datDesde.value, db)
    
    impresion_service.printReport Me.crpClientes, "LCtaCteClientesMy", db.sconection, , Array(Array("pCliID", Me.cboClienteD.ItemData(Me.cboClienteD.ListIndex)), Array("pDesde", toReportDate(Me.datDesde.value)), Array("pHasta", toReportDate(Me.datHasta.value)), Array("pSaldoDeudor", a2Decimales(curSaldoDeudor)))
    
    Me.MousePointer = 0

    Me.cmdConsultar.Enabled = True

End Sub

Private Sub cmdLista_Click()
Dim impresion_service As New clsCtlImpresion

    If Me.cboClienteD.ListIndex < 0 Then Exit Sub

    Me.cmdLista.Enabled = False
    
    Me.MousePointer = 11
    
    impresion_service.printReport Me.crpClientes, "rptCompCliente", db.sconection, , Array(Array("pCliID", Me.cboClienteD.ItemData(Me.cboClienteD.ListIndex)), Array("pDesde", toReportDate(Me.datDesde.value)), Array("pHasta", toReportDate(Me.datHasta.value)))

    Me.MousePointer = 0

    Me.cmdLista.Enabled = True

End Sub

Private Sub Form_Load()

    Me.datDesde.value = Date
    Me.datHasta.value = Date
    objCliente.fillCombo Me.cboClienteD, db
    
End Sub

Private Sub txtDesde_GotFocus()

    marcarseleccion Me.txtDesde
    
End Sub

Private Sub txtDesde_LostFocus()

On Error Resume Next

    objCliente.findByCodigo Val(Me.txtDesde.Text), db
    Me.cboClienteD.Text = objCliente.razon
    
End Sub
