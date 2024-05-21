VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ConVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ventas Mensual"
   ClientHeight    =   1725
   ClientLeft      =   1065
   ClientTop       =   1635
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4110
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   3360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtAnho 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboMes 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox chkDetallado 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Mes"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "ConVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
Dim impresion_service As New clsCtlImpresion

    If Me.cboMes.ListIndex < 0 Then Exit Sub
    
    If Val(Me.txtAnho.Text) < 2000 Then Exit Sub
    
    Me.cmdConsultar.Enabled = False
    
    Me.MousePointer = 11
    
    impresion_service.printReport Me.crpConsulta, "rptVentaMensual", db.sconection, , Array(Array("pDesde", toReportDate(primerDiaMes(Me.cboMes.ItemData(Me.cboMes.ListIndex), Val(Me.txtAnho.Text)))), Array("pHasta", toReportDate(ultimoDiaMes(Me.cboMes.ItemData(Me.cboMes.ListIndex), Val(Me.txtAnho.Text)))), Array("pDetalle", Me.chkDetallado.value))
    
    Me.MousePointer = 0

    Me.cmdConsultar.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
Dim strMeses As Variant

Dim intCiclo As Integer

    Me.cboMes.Clear

    strMeses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    
    For intCiclo = LBound(strMeses) To UBound(strMeses)
        Me.cboMes.AddItem strMeses(intCiclo)
        Me.cboMes.ItemData(Me.cboMes.NewIndex) = intCiclo - LBound(strMeses) + 1
    Next intCiclo
    
    Me.cboMes.ListIndex = 0
    
    Me.txtAnho.Text = Year(Date)
    
End Sub

Private Sub txtAnho_GotFocus()

    marcarseleccion Me.txtAnho
    
End Sub
