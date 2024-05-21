VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Consolidado"
   ClientHeight    =   1725
   ClientLeft      =   2385
   ClientTop       =   2040
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4110
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Hasta"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Desde"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
Dim impresion_service As New clsCtlImpresion

    Me.MousePointer = 11
    
    Me.cmdConsultar.Enabled = False
    
    impresion_service.printReport Me.crpConsulta, "rptConsolidado", db.sconection, Array("sValores", "sArticulos"), Array(Array("pDesde", toReportDate(Me.datDesde.value)), Array("pHasta", toReportDate(Me.datHasta.value)))
    
    Me.cmdConsultar.Enabled = True
    
    Me.MousePointer = 0

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
    
    Me.datDesde.value = Date
    Me.datHasta.value = Date

End Sub
