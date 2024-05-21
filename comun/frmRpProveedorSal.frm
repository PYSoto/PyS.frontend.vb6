VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRpProveedorSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos Cuentas Corrientes Proveedores"
   ClientHeight    =   1575
   ClientLeft      =   2430
   ClientTop       =   1950
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkSinSaldo 
      Caption         =   "Sin Saldo Anterior"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpProveedor 
      Left            =   3480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   3997697
      CurrentDate     =   39053
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   3997697
      CurrentDate     =   39053
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmRpProveedorSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.Value = Me.dtpDesde.Value
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cmdConsultar_Click()
Dim proveedor As New clsDAOProveedor

Dim ctlImp As New clsCtlImpresion
Dim ctlCom As New clsCtlCompras
    
    Me.cmdConsultar.Enabled = False
    Me.MousePointer = 11
    
    For Each proveedor In proveedor.collectionAll(db)
        ctlCom.updateSaldoCCTope proveedor.proveedorID, Me.dtpDesde.Value, db
        ctlCom.updateSaldoCCTope proveedor.proveedorID, Me.dtpHasta.Value + 1, db
    Next
    
    ctlImp.printReport Me.crpProveedor, "rptProveedorSaldo", db.sconection, , Array(Array("desde", toReportDate(Me.dtpDesde.Value)), Array("hasta", toReportDate(Me.dtpHasta.Value)), Array("sinSaldoAnterior", Me.chkSinSaldo.Value))
    
    Me.cmdConsultar.Enabled = True
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()

    Me.dtpDesde.Value = Date
    Me.dtpHasta.Value = Date
    
End Sub

