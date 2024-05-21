VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRpProvLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Proveedores"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4110
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRpProvLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
Dim ctlImp As New clsCtlImpresion
    
    Me.cmdImprimir.Enabled = False
    
    Me.MousePointer = 11
    
    ctlImp.printReport Me.crpConsulta, "rptProvLista", db.sconection
    
    Me.MousePointer = 0
    
    Me.cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub



