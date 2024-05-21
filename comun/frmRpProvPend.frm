VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRpProvPend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Pendientes"
   ClientHeight    =   1680
   ClientLeft      =   2385
   ClientTop       =   2040
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6000
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   9175041
      CurrentDate     =   39053
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   2760
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   9175041
      CurrentDate     =   39053
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmRpProvPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
Dim ctlImp As New clsCtlImpresion

    Me.cmdConsultar.Enabled = False

    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpConsulta, "rptProvPend", db.sconection, , Array(Array("desde", toReportDate(Me.dtpDesde.Value)), Array("hasta", toReportDate(Me.dtpHasta.Value)))
    
    Me.MousePointer = 0

    Me.cmdConsultar.Enabled = True
    
End Sub

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.Value = Me.dtpDesde.Value
    
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.dtpDesde.Value = Date
    Me.dtpHasta.Value = Date
    
End Sub



