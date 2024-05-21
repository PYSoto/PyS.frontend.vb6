VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form frmCopiar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Artículos"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6030
   Begin ComctlLib.ProgressBar prbCopia 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "Copiar"
      Height          =   372
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin ComctlLib.StatusBar sbrEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmCopiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopiar_Click()
Dim objArt As New clsDAOArticulos

Dim objArtAlias As New clsDAOArticulosAlias

Dim objArtAlter As New clsDAOArticulosAlter

Dim objPrv As New clsDAOProveedor

Dim colElem As Collection

    Me.MousePointer = 11
    Me.cmdCopiar.Enabled = False
    
    Me.prbCopia.Min = 0
    
    Me.sbrEstado.SimpleText = "Copiando ARTICULOS . . ."
    Me.sbrEstado.Refresh
    Set colElem = objArt.collectionAll
    Me.prbCopia.Value = Me.prbCopia.Min
    Me.prbCopia.Max = colElem.Count + 1
    For Each objArt In colElem
        DoEvents
        Me.prbCopia.Value = Me.prbCopia.Value + 1
        gArtMirror.saveArticulo objArt, db
    Next
    
    Me.sbrEstado.SimpleText = "Copiando ARTICULOSALIAS . . ."
    Me.sbrEstado.Refresh
    Set colElem = objArtAlias.collectionAll
    Me.prbCopia.Value = Me.prbCopia.Min
    Me.prbCopia.Max = colElem.Count + 1
    For Each objArtAlias In colElem
        DoEvents
        Me.prbCopia.Value = Me.prbCopia.Value + 1
        gArtMirror.saveArticuloAlias objArtAlias, db
    Next
    
    Me.sbrEstado.SimpleText = "Copiando ARTICULOSALTERNATIVO . . ."
    Me.sbrEstado.Refresh
    Set colElem = objArtAlter.collectionAll
    Me.prbCopia.Value = Me.prbCopia.Min
    Me.prbCopia.Max = colElem.Count + 1
    For Each objArtAlter In colElem
        DoEvents
        Me.prbCopia.Value = Me.prbCopia.Value + 1
        gArtMirror.saveArticuloAlter objArtAlter, db
    Next
    
    Me.sbrEstado.SimpleText = "Copiando PROVEEDORES . . ."
    Me.sbrEstado.Refresh
    Set colElem = objPrv.collectionAll(db)
    Me.prbCopia.Value = Me.prbCopia.Min
    Me.prbCopia.Max = colElem.Count + 1
    For Each objPrv In colElem
        DoEvents
        Me.prbCopia.Value = Me.prbCopia.Value + 1
        gArtMirror.saveProveedor objPrv, db
    Next
    
    Me.sbrEstado.SimpleText = ". . . Copia TERMINADA"
    
    Me.cmdCopiar.Enabled = True
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

