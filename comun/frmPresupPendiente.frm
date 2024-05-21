VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPresupPendiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos Pendientes"
   ClientHeight    =   4215
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   11790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFacturados 
      Caption         =   "Facturados"
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboDias 
      Height          =   315
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   9840
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdPend 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5106
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dias"
      Height          =   195
      Left            =   9840
      TabIndex        =   3
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmPresupPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public preID As Variant

Private Sub cboDias_Click()

    Form_Activate
    
End Sub

Private Sub chkFacturados_Click()

    Form_Activate
    
End Sub

Private Sub cmdSalir_Click()
    
    Me.Hide
    
End Sub

Private Sub Form_Activate()
Dim objPr As New clsDAOPresupuesto

Dim objCl As New clsDAOClientes

Dim objMC As New clsDAOMovclie

Dim strFac As String

    Me.grdPend.Rows = 1
    Me.grdPend.Redraw = False
    For Each objPr In objPr.collectionPendientes(Me.cboDias.ItemData(Me.cboDias.ListIndex), db, Me.chkFacturados.Value)
        strFac = ""
        If objCl.codigo <> objPr.cliID Then objCl.findByCodigo objPr.cliID, db
        With objPr
            If .clavemovclie > 0 Then
                objMC.clave = .clavemovclie
                objMC.findByPrimaryKey db
                
                strFac = "Fac " & objMC.tipocompro & " " & Format(objMC.prefijo, "0000") & "-" & Format(objMC.nroComprob, "00000000")
            End If
            Me.grdPend.AddItem modGrid.array2itemGrid(Array(.preID, objCl.razon, .fecha, .fechavto, .observac, "", strFac))
            Me.grdPend.RowData(Me.grdPend.Rows - 1) = .preID
            modGrid.setCheckGrid Me.grdPend, Me.grdPend.Rows - 1, 5, IIf(.ctacte = 1, True, False)
        End With
    Next
    Me.grdPend.Redraw = True
    
End Sub

Private Sub Form_Load()
    
    modGrid.makeGrid2 Me.grdPend, Array(Array("Pre", 600), Array("Cliente", 2000), Array("Fecha", 1200), Array("Vencimiento", 1200), Array("Observaciones", 3700), Array("CC", 300), Array("Factura", 1900)), 0, 1, flexSelectionByRow
    
    Me.cboDias.AddItem "7 días"
    Me.cboDias.ItemData(0) = 7
    Me.cboDias.AddItem "15 días"
    Me.cboDias.ItemData(1) = 15
    Me.cboDias.AddItem "30 días"
    Me.cboDias.ItemData(2) = 30
    Me.cboDias.AddItem "60 días"
    Me.cboDias.ItemData(3) = 60
    Me.cboDias.AddItem "Todos"
    Me.cboDias.ItemData(4) = 1000
    Me.cboDias.ListIndex = 0
    
    preID = Null

End Sub

Private Sub grdPend_DblClick()

    If Me.grdPend.Row < 1 Then Exit Sub
    
    preID = Me.grdPend.RowData(Me.grdPend.Row)
    
    cmdSalir_Click
    
End Sub
