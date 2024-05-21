VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListaRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   2430
   ClientLeft      =   2430
   ClientTop       =   1950
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5985
   Begin VB.TextBox txtFin 
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtInicio 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkNumero 
      Caption         =   "por Número"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   "por Fecha"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkCliente 
      Caption         =   "por Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin Crystal.CrystalReport crpClientes 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   39053
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fin"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Inicio"
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta Fecha :"
      Height          =   195
      Left            =   4080
      TabIndex        =   6
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde Fecha :"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   1050
   End
End
Attribute VB_Name = "frmListaRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objCliente As New clsDAOClientes

Private objMovclie As New clsDAOMovclie

Private Sub chkFecha_Click()
Dim datDesde As Date
Dim datHasta As Date

    If Me.chkFecha.value = CheckBoxConstants.vbChecked Then
        Me.datDesde.Enabled = True
        Me.datHasta.Enabled = True
    Else
        objMovclie.extremosFechaComprob datDesde, datHasta, db
        Me.datDesde.Enabled = False
        Me.datDesde.value = datDesde
        Me.datHasta.Enabled = False
        Me.datHasta.value = datHasta
    End If
    
End Sub

Private Sub chkNumero_Click()
Dim lngInicio As Long
Dim lngFin As Long

    If Me.chkNumero.value = CheckBoxConstants.vbChecked Then
        Me.txtInicio.Locked = False
        Me.txtFin.Locked = False
    Else
        objMovclie.extremosNroComprob lngInicio, lngFin, db
        Me.txtInicio.Locked = True
        Me.txtInicio.Text = lngInicio
        Me.txtFin.Locked = True
        Me.txtFin.Text = lngFin
    End If
    
End Sub

Private Sub cmdConsultar_Click()
Dim lngCliDesde As Long
Dim lngCliHasta As Long

Dim impresion_service As New clsCtlImpresion

    If Me.cboCliente.ListIndex < 0 Then Exit Sub
    
    Me.cmdConsultar.Enabled = False

    objCliente.extremos lngCliDesde, lngCliHasta, db
    
    If Me.chkCliente.value = CheckBoxConstants.vbChecked Then
        lngCliDesde = Me.cboCliente.ItemData(Me.cboCliente.ListIndex)
        lngCliHasta = lngCliDesde
    End If
    
    Me.MousePointer = 11
    
    impresion_service.printReport Me.crpClientes, "LRecibosVentas", db.sconection, , Array(Array("pCliDesde", lngCliDesde), Array("pCliHasta", lngCliHasta), Array("pDesde", toReportDate(Me.datDesde.value)), Array("pHasta", toReportDate(Me.datHasta.value)), Array("pNroDesde", Val(Me.txtInicio.Text)), Array("pNroHasta", Val(Me.txtFin.Text))), Array("+{movclie.cgoclie}", IIf(Me.chkNumero.value = CheckBoxConstants.vbChecked, "+{movclie.nrocomprob}", "+{movclie.fechacomprob}"))
        
    Me.MousePointer = 0
    
    Me.cmdConsultar.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    objCliente.fillCombo Me.cboCliente, db
    
    Call chkNumero_Click
    
    Call chkFecha_Click
    
End Sub

Private Sub txtFin_GotFocus()

    marcarseleccion Me.txtFin
    
End Sub

Private Sub txtInicio_GotFocus()

    marcarseleccion Me.txtInicio
    
End Sub
