VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiCompras 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo de Compras"
   ClientHeight    =   7575
   ClientLeft      =   810
   ClientTop       =   2610
   ClientWidth     =   10305
   LinkTopic       =   "MDIForm1"
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sokAplicacion 
      Left            =   840
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mCompra 
      Caption         =   "Compras"
      Begin VB.Menu mComprobante 
         Caption         =   "Registro de Comprobantes"
      End
      Begin VB.Menu mProvPagos 
         Caption         =   "Registro de Pagos"
      End
      Begin VB.Menu ray01 
         Caption         =   "-"
      End
      Begin VB.Menu mAnulacion 
         Caption         =   "Reimpresión y Anulación de Comprobantes"
      End
   End
   Begin VB.Menu mConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu mPendientes 
         Caption         =   "Facturas Pendientes"
      End
      Begin VB.Menu linea01 
         Caption         =   "-"
      End
      Begin VB.Menu mProveedorCC 
         Caption         =   "Cuenta Corriente por Proveedor"
      End
      Begin VB.Menu mSaldos 
         Caption         =   "Saldos Cuentas Corrientes"
      End
      Begin VB.Menu mRpProveedor 
         Caption         =   "Proveedores"
      End
   End
   Begin VB.Menu mParametro 
      Caption         =   "Parámetros"
      Begin VB.Menu mArticulo 
         Caption         =   "Artículos"
      End
      Begin VB.Menu mProveedor 
         Caption         =   "P&roveedores"
      End
      Begin VB.Menu ray2727 
         Caption         =   "-"
      End
      Begin VB.Menu mAumentarPreciosProveedor 
         Caption         =   "Cambiar Precios por Proveedor"
      End
   End
   Begin VB.Menu mVentana 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mHorizontal 
         Caption         =   "Mosaico Horizontal"
      End
      Begin VB.Menu mVertical 
         Caption         =   "Mosaico Vertical"
      End
   End
   Begin VB.Menu mFin 
      Caption         =   "Fin"
   End
End
Attribute VB_Name = "mdiCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private usuarioactivo As New clsDAOUsuarioActivo

Private Sub mAnulacion_Click()

    frmProveedorAnul.Show
    
End Sub

Private Sub mArticulo_Click()

    frmArticulos.Show
    
End Sub

Private Sub mAumentarPreciosProveedor_Click()

    frmAumentarPrecioProveedor.Show
    
End Sub

Private Sub mComprobante_Click()

    frmProvCompras.Show
    
End Sub

Private Sub MDIForm_Activate()

    gEmpresa.findLast db
    mdiCompras.Caption = cntArchivo & " - Módulo de Compras - " & gEmpresa.nombre & " - " & gEmpresa.ubicacion
    
    If Not gUsuario.hayUsuarios(db) Then Exit Sub
    
    If Not usuarioactivo.hayUsuarioActivo(Me.sokAplicacion.LocalIP, Me.hwnd, db) Then frmLogin.Show vbModal
    
    mostrar True
    
End Sub

Private Sub MDIForm_Load()

    Call ponerConfiguracionRegional

    Set gMDI = Me
    
    gCon.configureDB
    
    usuarioactivo.eliminarUsuarioActivo Me.sokAplicacion.LocalIP, Me.hwnd, db
    
    mostrar False
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    End
    
End Sub

Public Sub mostrar(ByVal pMostrar As Boolean)

    mdiCompras.mCompra.Visible = pMostrar
    mdiCompras.mParametro.Visible = pMostrar
    mdiCompras.mVentana.Visible = pMostrar
    mdiCompras.mConsultas.Visible = pMostrar
    mdiCompras.mFin.Visible = pMostrar
    
End Sub

Private Sub mFin_Click()

    db.closeDB
    End
    
End Sub

Private Sub mPendientes_Click()

    frmRpProvPend.Show
    
End Sub

Private Sub mProveedor_Click()

    frmProveedor.Show
    
End Sub

Private Sub mProveedorCC_Click()

    frmRpProveedorCC.Show
    
End Sub

Private Sub mProvPagos_Click()

    frmProvPagos.Show
    
End Sub

Private Sub mRpProveedor_Click()

    frmRpProvLista.Show
    
End Sub

Private Sub mSaldos_Click()

    frmRpProveedorSal.Show
    
End Sub
