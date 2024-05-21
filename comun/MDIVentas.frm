VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiVentas 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo Ventas"
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
      Begin VB.Menu AbmClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu AbmArtic 
         Caption         =   "Artículos"
      End
      Begin VB.Menu lin 
         Caption         =   "-"
      End
      Begin VB.Menu sale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu movim 
      Caption         =   "Movimientos "
      Begin VB.Menu Factur 
         Caption         =   "Facturación"
      End
      Begin VB.Menu Rec 
         Caption         =   "Recibos"
      End
   End
   Begin VB.Menu consulta 
      Caption         =   "Consultas"
      Begin VB.Menu FacPen 
         Caption         =   "Facturas Pendientes"
      End
      Begin VB.Menu ConFactu 
         Caption         =   "Facturación"
      End
      Begin VB.Menu LCtaCte 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu mSaldosClientes 
         Caption         =   "Saldos Cuentas Corrientes"
      End
      Begin VB.Menu LRecibos 
         Caption         =   "Recibos"
      End
      Begin VB.Menu raya2 
         Caption         =   "-"
      End
      Begin VB.Menu mConsolidado 
         Caption         =   "Consolidado"
      End
   End
   Begin VB.Menu bd 
      Caption         =   "Mantenimientos"
      Begin VB.Menu Anula 
         Caption         =   "Anulaciones"
      End
      Begin VB.Menu ModIva 
         Caption         =   "Alícuotas de I.V.A."
      End
      Begin VB.Menu mImportar 
         Caption         =   "Importar Lista IVECO"
      End
      Begin VB.Menu mExportar 
         Caption         =   "Exportar/Importar San Juan"
      End
      Begin VB.Menu mCopiaArt 
         Caption         =   "Copiar Artículos"
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
Attribute VB_Name = "mdiVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objUsuarioActivo As New clsDAOUsuarioActivo

Private Sub ABMArtic_Click()

    frmArticulos.Show
    
End Sub

Private Sub AbmClientes_Click()

    Clientes.Show
    
End Sub

Private Sub Anula_Click()
    
    MovVentasAnula.Show
    
End Sub

Private Sub ConFactu_Click()
    
    ConVentas.Show
    
End Sub

Private Sub FacPen_Click()

    frmFacPend.Show
    
End Sub

Private Sub Factur_Click()

    MovVentas.Show
    
End Sub

Private Sub LCtaCte_Click()

    LCtaCtecli.Show
    
End Sub

Private Sub LRecibos_Click()

    frmListaRecibos.Show
    
End Sub

Private Sub mConsolidado_Click()

    frmConsolidado.Show
    
End Sub

Private Sub mCopiaArt_Click()

    frmCopiar.Show
    
End Sub

Private Sub MDIForm_Activate()

    Set gMDI = mdiVentas

    gEmpresa.findLast db
    mdiVentas.Caption = cntArchivo & " - Módulo de Ventas - " & gEmpresa.nombre & " - " & gEmpresa.ubicacion
    
    If Not gUsuario.hayUsuarios(db) Then Exit Sub
    
    If Not objUsuarioActivo.hayUsuarioActivo(Me.sokAplicacion.LocalIP, Me.hwnd, db) Then frmLogin.Show vbModal
    
    mostrar True
    
End Sub

Private Sub MDIForm_Load()

    gCon.configureDB

    ponerConfiguracionRegional

    objUsuarioActivo.eliminarUsuarioActivo Me.sokAplicacion.LocalIP, Me.hwnd, db
    
    mostrar False
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    End
    
End Sub

Private Sub mExportar_Click()

    frmExportar.Show
    
End Sub

Private Sub mImportar_Click()

    frmImportarIVECO.Show
    
End Sub

Private Sub ModIva_Click()

    ModifIva.Show
    
End Sub

Private Sub mSaldosClientes_Click()

    frmSaldosClientes.Show
    
End Sub

Private Sub Rec_Click()

    Recibos.Show
    
End Sub

Private Sub sale_Click()

    End
    
End Sub

Public Sub mostrar(ByVal pMostrar As Boolean)

    mdiVentas.mae.Visible = pMostrar
    mdiVentas.movim.Visible = pMostrar
    mdiVentas.consulta.Visible = pMostrar
    mdiVentas.bd.Visible = pMostrar
    mdiVentas.venta.Visible = pMostrar
    
End Sub


