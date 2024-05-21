VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSaldosClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos Cuentas Corrientes Clientes"
   ClientHeight    =   1710
   ClientLeft      =   2430
   ClientTop       =   1950
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4095
   Begin VB.CheckBox chkSinSaldo 
      Caption         =   "Sin Saldo Anterior"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpClientes 
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmSaldosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objMovclie As New clsDAOMovclie

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub cmdConsultar_Click()
Dim objTMB As New clsDAOTMBalance

Dim impresion_service As New clsCtlImpresion
    
    Me.cmdConsultar.Enabled = False
    
    Me.MousePointer = 11
    
    If Me.chkSinSaldo.value = 1 Then
        objTMB.saldosImportesSinAnterior Me.hWnd, Me.datDesde.value, Me.datHasta.value, db
    Else
        objTMB.saldosImportesConAnterior Me.hWnd, Me.datDesde.value, Me.datHasta.value, db
    End If
    
    impresion_service.printReport Me.crpClientes, "rptSaldosClientes", db.sconection, , Array(Array("pDesde", toReportDate(Me.datDesde.value)), Array("pHasta", toReportDate(Me.datHasta.value)), Array("phWnd", Me.hWnd))
    
    With objTMB
        .hWnd = Me.hWnd
        
        .deleteAll db
    End With
    
    Me.MousePointer = 0

    Me.cmdConsultar.Enabled = True

End Sub

Private Sub Form_Load()

    Me.datDesde.value = Date
    Me.datHasta.value = Date
    
End Sub

