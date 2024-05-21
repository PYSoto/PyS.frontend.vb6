VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConArtic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Ventas por Artículos"
   ClientHeight    =   2070
   ClientLeft      =   2385
   ClientTop       =   2040
   ClientWidth     =   3885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3885
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   49545217
      CurrentDate     =   39053
   End
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   49545217
      CurrentDate     =   39053
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta Fecha :"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde Fecha :"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "ConArtic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()

On Error Resume Next

    With Me.crpConsulta
        .Reset
        .ReportFileName = App.Path & IIf(Me.chkDetallado.Value = 1, "\VtasDetArtMy.rpt", "\VtasResArtMy.rpt")
        .Connect = DB.SConection
        .ParameterFields(0) = "pEmpresa;" & gEmpresa.nombre & ";TRUE"
        .ParameterFields(1) = "pDesde;" & toReportDate(Me.datDesde.Value) & ";TRUE"
        .ParameterFields(2) = "pHasta;" & toReportDate(Me.datHasta.Value) & ";TRUE"
        
        frmImpresora.Show vbModal
        frmImpresora.cargar Me.crpConsulta
        If Not frmImpresora.Cancel Then .Action = 1
    
    End With

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Activate()

    Me.datDesde.Value = Date
    Me.datHasta.Value = Date

End Sub
