VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListaPres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Presupuestos por Período"
   ClientHeight    =   1410
   ClientLeft      =   3270
   ClientTop       =   1845
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin Crystal.CrystalReport crpReporte 
      Left            =   1080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1695
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
      Format          =   49610753
      CurrentDate     =   39059
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49610753
      CurrentDate     =   39059
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmListaPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
Dim ctlImp As New clsCtlImpresion

    Me.cmdConsultar.Enabled = False
    
    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpReporte, "rptPresPeriodo", db.sconection, , Array(Array("pDesde", toReportDate(Me.datDesde.Value)), Array("pHasta", toReportDate(Me.datHasta.Value)))
    
    Me.MousePointer = 0
    
    Me.cmdConsultar.Enabled = True

End Sub

Private Sub cmdDuplicar_Click()

    Me.datHasta.Value = Me.datDesde.Value
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.datDesde.Value = Date
    Me.datHasta.Value = Date

End Sub

