VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDepClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depurar Clientes"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6015
   Begin ComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDepurar 
      Caption         =   "Depurar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmDepClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDepurar_Click()
Dim objCli As New clsDAOClientes

Dim colCli As Collection

Dim ctlMir As New clsCtlMirror

On Error Resume Next

    Me.cmdDepurar.Enabled = False
    
    Me.MousePointer = 11
    
    Set colCli = objCli.collectionAll(db)
    
    Me.prbProgreso.Min = 0
    Me.prbProgreso.Max = colCli.Count
    Me.prbProgreso.Value = Me.prbProgreso.Min
    
    For Each objCli In colCli
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        ctlMir.depuraCliente objCli, db
    Next
    
    Me.MousePointer = 0
    
    Me.cmdDepurar.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub
