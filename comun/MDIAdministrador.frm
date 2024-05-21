VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiAdministrador 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo Presupuesto"
   ClientHeight    =   5565
   ClientLeft      =   1410
   ClientTop       =   2385
   ClientWidth     =   7410
   LinkTopic       =   "MDIForm1"
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sokAplicacion 
      Left            =   600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mae 
      Caption         =   "Maestros"
      Begin VB.Menu mArticulos 
         Caption         =   "Artículos"
      End
      Begin VB.Menu mEmpresa 
         Caption         =   "Empresa"
      End
      Begin VB.Menu ray72727 
         Caption         =   "-"
      End
      Begin VB.Menu sale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu venta 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mh 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu mv 
         Caption         =   "Mosaico Vertical"
      End
   End
End
Attribute VB_Name = "mdiAdministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objUsuarioActivo As New clsDAOUsuarioActivo

Private Sub mArticulos_Click()
    
    frmArticulos.Show
    
End Sub

Private Sub MDIForm_Activate()

    Set gMDI = MDIAdministrador

    gEmpresa.findLast db
    Me.Caption = cntArchivo & " - Módulo de Administrador - " & gEmpresa.nombre & " - " & gEmpresa.ubicacion
    
    If Not gUsuario.hayUsuarios(db) Then Exit Sub
    
    If Not objUsuarioActivo.hayUsuarioActivo(Me.sokAplicacion.LocalIP, Me.hwnd, db) Then frmLogin.Show vbModal
    
    mostrar True
    
End Sub

Private Sub MDIForm_Load()

    ponerConfiguracionRegional
    
    gCon.configureDB

    objUsuarioActivo.eliminarUsuarioActivo Me.sokAplicacion.LocalIP, Me.hwnd, db
    
    mostrar False
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    End
    
End Sub

Private Sub mEmpresa_Click()

    frmEmpresa.Show
    
End Sub

Private Sub sale_Click()

    End
    
End Sub

Public Sub mostrar(ByVal pMostrar As Boolean)

    Me.mae.Visible = pMostrar
    Me.venta.Visible = pMostrar
    
End Sub


