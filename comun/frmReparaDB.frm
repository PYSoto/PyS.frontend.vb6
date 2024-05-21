VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmReparaDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparar Base de Datos"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblEstado 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmReparaDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    gCon.configureDB
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = 100
    Me.prbProgreso = Me.prbProgreso.Min
    
End Sub

Private Sub Form_Activate()
Dim rstQuery As Recordset

Dim strSQL As String

Dim intCantidad As Integer

    Me.MousePointer = 11
    
    ' Cuenta registros
    
    DB.openDB
    
    Set rstQuery = DB.activa.OpenSchema(adSchemaTables)
    
    intCantidad = 0
    
    Do While Not rstQuery.EOF
        
        intCantidad = intCantidad + 1
        
        rstQuery.MoveNext
    
    Loop
    
    rstQuery.Close
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = intCantidad + IIf(intCantidad > 1, 0, 1)
    
    intCantidad = 0
    ' Repara tablas
    
    Set rstQuery = DB.activa.OpenSchema(adSchemaTables)
    
    Do While Not rstQuery.EOF
        
        intCantidad = intCantidad + 1
        
        Me.prbProgreso.Value = intCantidad
        
        DoEvents
        
        Me.lblEstado.Caption = "Reparando TABLA '" & rstQuery!TABLE_NAME & "' . . ."
        Me.lblEstado.Refresh
        
        strSQL = "REPAIR TABLE " & rstQuery!TABLE_NAME & ";"
        DB.activa.Execute strSQL
        
        rstQuery.MoveNext
    
    Loop
    
    rstQuery.Close
    
    Me.lblEstado.Caption = " . . . Reparación Terminada"
    MsgBox "Reparación TERMINADA . . ."
    
    Me.MousePointer = 0
    
    DB.closeDB
    
    End

End Sub
