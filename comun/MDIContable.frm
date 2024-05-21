VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiContable 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Contable"
   ClientHeight    =   5565
   ClientLeft      =   1935
   ClientTop       =   2115
   ClientWidth     =   7410
   LinkTopic       =   "MDIForm1"
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sokAplicacion 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2760
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mae 
      Caption         =   "Maestros"
      Begin VB.Menu mClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu ray3737 
         Caption         =   "-"
      End
      Begin VB.Menu sale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mMovimientos 
      Caption         =   "Movimientos"
      Begin VB.Menu mVentaManual 
         Caption         =   "Facturación Manual"
      End
      Begin VB.Menu mMovClieSAC 
         Caption         =   "Ventas de Otro Negocio"
      End
   End
   Begin VB.Menu Consulta 
      Caption         =   "Consultas"
      Begin VB.Menu mIvaCompras 
         Caption         =   "Libro de Iva Compras"
      End
      Begin VB.Menu LivaVenta 
         Caption         =   "Libro de Iva Ventas"
      End
   End
   Begin VB.Menu bd 
      Caption         =   "Mantenimientos"
      Begin VB.Menu mCambiarClave 
         Caption         =   "Cambiar contraseña"
      End
   End
   Begin VB.Menu venta 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mh 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu mV 
         Caption         =   "Mosaico Vertical"
      End
   End
End
Attribute VB_Name = "mdiContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objUsuarioActivo As New clsDAOUsuarioActivo

Private Sub LivaVenta_Click()

    LIvaVentas.Show
    
End Sub

Private Sub mCambiarClave_Click()

    frmCambiarClave.Show
    
End Sub

Private Sub mClientes_Click()

    Clientes.Show
    
End Sub

Private Sub MDIForm_Activate()

    Set gMDI = MDIContable

    gEmpresa.findLast db
    MDIContable.Caption = cntArchivo & " - Módulo Contable - " & gEmpresa.nombre & " - " & gEmpresa.ubicacion
    
    If Not gUsuario.hayUsuarios(db) Then Exit Sub
    
    If Not objUsuarioActivo.hayUsuarioActivo(Me.sokAplicacion.LocalIP, Me.hwnd, db) Then frmLogin.Show vbModal
    
    mostrar True
    
End Sub

Private Sub MDIForm_Load()

    Call ponerConfiguracionRegional
    
    gCon.configureDB
    
    objUsuarioActivo.eliminarUsuarioActivo Me.sokAplicacion.LocalIP, Me.hwnd, db
    
    mostrar False
    
End Sub

Private Sub mIvaCompras_Click()

    LIvaCompras.Show
    
End Sub

Private Sub mMovClieSAC_Click()

    frmMovClieSAC.Show
    
End Sub

Private Sub mVentaManual_Click()

    frmVentaManual.Show
    
End Sub

Private Sub sale_Click()

    End
    
End Sub

Public Sub mostrar(ByVal pMostrar As Boolean)

    MDIContable.mae.Visible = pMostrar
    MDIContable.mMovimientos.Visible = pMostrar
    MDIContable.Consulta.Visible = pMostrar
    MDIContable.bd.Visible = pMostrar
    MDIContable.venta.Visible = pMostrar
    
End Sub



