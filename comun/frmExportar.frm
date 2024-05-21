VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   372
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdlArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   ". . ."
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin ComctlLib.StatusBar sbrEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1185
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
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()

    With Me.cdlArchivo
        .Filter = "San Juan|*.sje"
        
        .ShowSave
        
        Me.txtArchivo.Text = .FileName
    End With
    
End Sub

Private Sub cmdExportar_Click()
Dim objArticulo As New clsDAOArticulos

Dim objArticulosAlias As New clsDAOArticulosAlias

Dim objArticulosAlternativo As New clsDAOArticulosAlter

Dim objProveedor As New clsDAOProveedor

    If Me.txtArchivo.Text = "" Then Exit Sub
    
    ' Borra el archivo
    Open Me.txtArchivo.Text For Output As #1
    Close #1
    
    Me.MousePointer = 11
    
    Me.sbrEstado.SimpleText = "Exportando ARTICULOS . . ."
    Me.sbrEstado.Refresh
    objArticulo.exportar Me.txtArchivo.Text, db
    
    Me.sbrEstado.SimpleText = "Exportando ARTICULOSALIAS . . ."
    Me.sbrEstado.Refresh
    objArticulosAlias.exportar Me.txtArchivo.Text, db
    
    Me.sbrEstado.SimpleText = "Exportando ARTICULOSALTERNATIVO . . ."
    Me.sbrEstado.Refresh
    objArticulosAlternativo.exportar Me.txtArchivo.Text, db
    
    Me.sbrEstado.SimpleText = "Exportando PROVEEDORES . . ."
    Me.sbrEstado.Refresh
    objProveedor.exportar Me.txtArchivo.Text, db
    
    Me.sbrEstado.SimpleText = ". . . Exportación TERMINADA"
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdImportar_Click()
Dim strLinea As String
Dim strLineaIntermedia As String

Dim arrValores As Variant

Dim objArticulo As New clsDAOArticulos

Dim objArticulosAlias As New clsDAOArticulosAlias

Dim objArticulosAlternativo As New clsDAOArticulosAlter

Dim objProveedor As New clsDAOProveedor

Dim lngContador As Long

On Error Resume Next

    objArticulosAlias.limpiar db
    
    objArticulosAlternativo.limpiar db
    
    objProveedor.truncate db
    
    lngContador = 0

    Open Me.txtArchivo.Text For Input As #1
    
    Line Input #1, strLinea
    
    Do While Not EOF(1)
        DoEvents
    
        lngContador = lngContador + 1
        
        Do While Not Right(strLinea, 1) = ";"
            Line Input #1, strLineaIntermedia
            strLinea = strLinea & Chr(10) & Chr(13) & strLineaIntermedia
        Loop
        
        Me.sbrEstado.SimpleText = lngContador & " - " & strLinea
        Me.sbrEstado.Refresh
        
        db.execute strLinea
        
        Line Input #1, strLinea
    Loop
    
    Close #1
    
    Me.sbrEstado.SimpleText = ". . . Importación TERMINADA"
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub txtArchivo_GotFocus()

    marcarseleccion Me.txtArchivo
    
End Sub
