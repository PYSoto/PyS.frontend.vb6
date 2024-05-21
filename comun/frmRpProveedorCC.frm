VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRpProveedorCC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente por Proveedor"
   ClientHeight    =   1590
   ClientLeft      =   2505
   ClientTop       =   1860
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   6000
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtProveedor 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkSinSaldo 
      Caption         =   "Sin Saldo Anterior"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39072
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   5400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   31260673
      CurrentDate     =   39072
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   465
   End
End
Attribute VB_Name = "frmRpProveedorCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private proveedor As New clsDAOProveedor

Private Sub cmdConsultar_Click()
Dim ctlImp As New clsCtlImpresion
Dim ctlCom As New clsCtlCompras

    If proveedor.proveedorID = 0 Then Exit Sub

    Me.cmdConsultar.Enabled = False
    
    If Not ctlCom.updateSaldoCCTope(proveedor.proveedorID, Me.dtpDesde.Value, db) Then MsgBox "ERROR: No se pudo actualizar el Saldo"

    ctlImp.printReport Me.crpConsulta, "rptProveedorCC", db.sconection, Array("valores"), Array(Array("proveedorID", proveedor.proveedorID), Array("desde", toReportDate(Me.dtpDesde.Value)), Array("hasta", toReportDate(Me.dtpHasta.Value)), Array("sinSaldoAnterior", Me.chkSinSaldo.Value))

    Me.cmdConsultar.Enabled = True

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.dtpDesde.Value = Date
    Me.dtpHasta.Value = Date

    Set proveedor = New clsDAOProveedor
    If proveedorglobal.proveedorID > 0 Then Set proveedor = proveedorglobal
    Me.txtProveedor.Text = proveedor.textSearch

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
    
End Sub
