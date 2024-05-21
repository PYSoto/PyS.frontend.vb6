VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegInfoCVSH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Régimen de Información de Compras y Ventas - SH"
   ClientHeight    =   2430
   ClientLeft      =   3270
   ClientTop       =   1845
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   7920
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2175
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   372
      Left            =   4080
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdArchivo 
      Height          =   255
      Left            =   7440
      Picture         =   "frmRegInfoCVSH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton cmdDuplicar 
      Caption         =   ">"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   372
      Left            =   6000
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   41418753
      CurrentDate     =   39059
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   41418753
      CurrentDate     =   39059
   End
   Begin MSComDlg.CommonDialog cdlArchivo 
      Left            =   4080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmRegInfoCVSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()
Dim strFile As String
Dim strPath As String
    
    Me.cdlArchivo.DialogTitle = "Buscar Carpeta"
    Me.cdlArchivo.ShowOpen
    
    strFile = modConv.parseFilename(Me.cdlArchivo.FileName, strPath)
    
    Me.txtArchivo.Text = strPath & "REGINFO_CV_CABECERA.txt"

End Sub

Private Sub cmdDuplicar_Click()

    Me.dtpHasta.Value = Me.dtpDesde.Value
    
End Sub

Private Sub cmdGenerar_Click()
Dim strFile As String
Dim strPath As String

Dim colDB As New Collection

    Me.MousePointer = 11
    
    colDB.add dbmdzsh
    
    strFile = modConv.parseFilename(Me.cdlArchivo.FileName, strPath)

    modRegInfCV.makeFile "30-57295624-1", strPath, Me.dtpDesde.Value, Me.dtpHasta.Value, colDB, Me.stbEstado, Me.prbProgreso
    
    Me.MousePointer = 0
    
    MsgBox "... Proceso TERMINADO", , "Régimen Información"
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.dtpDesde.Value = Date
    Me.dtpHasta.Value = Date
    
End Sub
