VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MovVentasAnula 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Comprobantes "
   ClientHeight    =   4110
   ClientLeft      =   915
   ClientTop       =   1635
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdCRecibo 
      Caption         =   "Imprimir C.Recibo"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datHasta 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100401153
      CurrentDate     =   39487
   End
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid grdComprobantes 
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker datDesde 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100401153
      CurrentDate     =   39487
   End
   Begin Crystal.CrystalReport crpConsulta 
      Left            =   7320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Comprobantes Ingresados"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1845
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "MovVentasAnula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clientemov As New clsDAOMovclie
Private cliente As New clsDAOClientes

Private Sub popupImputados(pClave As Long)
Dim clientecobro As New clsDAOFactcobradas
Dim clientemovlocal As New clsDAOMovclie
Dim comprobante As New clsDAOTiposComprob

    For Each clientecobro In clientecobro.collectionByClaveMovC(pClave)
        clientemovlocal.clave = clientecobro.clavepago
        clientemovlocal.findByPrimaryKey db
        comprobante.codigo = clientemovlocal.cgocomprob
        comprobante.findByPrimaryKey db
        
        MsgBox "Debe ANULAR -> " & comprobante.descripcion & " - " & Format(clientemovlocal.prefijo, "0000") & "-" & Format(clientemovlocal.nroComprob, "00000000")
    Next
    
End Sub

Private Sub cargarComprobantes()
Dim intRow As Integer

Dim comprobante As New clsDAOTiposComprob

On Error Resume Next

    If Me.cboCliente.ListIndex < 0 Then Exit Sub
    
    Me.grdComprobantes.Rows = 1
    
    For Each clientemov In clientemov.collectionAnulables(Me.cboCliente.ItemData(Me.cboCliente.ListIndex), Me.datDesde.value, Me.datHasta.value, db)
        With Me.grdComprobantes
            .Rows = .Rows + 1
            intRow = .Rows - 1
            .TextMatrix(intRow, 8) = clientemov.cgocomprob
            comprobante.negID = clientemov.negID
            comprobante.codigo = clientemov.cgocomprob
            comprobante.findByPrimaryKey db
            .TextMatrix(intRow, 1) = comprobante.comboText
            .TextMatrix(intRow, 2) = clientemov.prefijo
            .TextMatrix(intRow, 3) = clientemov.nroComprob
            .TextMatrix(intRow, 4) = clientemov.fechacomprob
            .TextMatrix(intRow, 6) = Format(clientemov.importe, "0.00")
            .TextMatrix(intRow, 7) = Format(clientemov.cancelado, "0.00")
            .TextMatrix(intRow, 0) = clientemov.clave
        End With
    Next
    
End Sub

Private Sub eliminar(ByVal pClaveMovClie As Long)
Dim comprobante As New clsDAOTiposComprob
Dim valormov As New clsDAOValores
Dim clientecobro As New clsDAOFactcobradas
Dim articulomov As New clsDAODetartic

    With clientemov
        .clave = pClaveMovClie
        .findByPrimaryKey db
    
        comprobante.codigo = clientemov.cgocomprob
        comprobante.findByPrimaryKey db
        
        If comprobante.factelect <> 0 Then
            MsgBox "ERROR: Comprobante ELECTRONICO"
            Exit Sub
        End If
    
        If comprobante.libroiva <> 0 Then
            .importe = 0
            .neto = 0
            .cancelado = 0
            .montoiva = 0
            .anulada = 1
            
            .save db
        Else
            .delete db
        End If
    End With
    
    For Each valormov In valormov.collectionByClaveMovClie(pClaveMovClie)
        valormov.delete
    Next
    
    For Each clientecobro In clientecobro.collectionByClavePago(pClaveMovClie)
        With clientemov
            .clave = clientecobro.clavemovc
            .findByPrimaryKey db
            
            .cancelado = .cancelado - clientecobro.importe
            
            .save db
        End With
        
        clientecobro.delete
    Next
    
    For Each articulomov In articulomov.collectionByClaveMovClie(pClaveMovClie)
        articulomov.delete
    Next
    
End Sub

Private Sub cmdAnular_Click()
Dim comprobante As New clsDAOTiposComprob

    If Me.cboCliente.ListIndex < 0 Then Exit Sub
    
    With Me.grdComprobantes
        If .row = 0 Then Exit Sub
        
        clientemov.clave = Val(.TextMatrix(.row, 0))
        clientemov.findByPrimaryKey db
        
        comprobante.codigo = clientemov.cgocomprob
        comprobante.findByPrimaryKey db
        
        If clientemov.cancelado <> 0 And comprobante.contado = 0 Then
            popupImputados clientemov.clave
            Exit Sub
        End If
        
        If MsgBox("Está Seguro que desea ANULAR ?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        
        Call eliminar(clientemov.clave)
    End With
        
    Call cargarComprobantes
        
End Sub

Private Sub cmdCRecibo_Click()
Dim impresion_service As New clsCtlImpresion

    If Me.grdComprobantes.row = 0 Then Exit Sub

    Me.cmdCRecibo.Enabled = False
    
    Me.MousePointer = 11
    
    impresion_service.printReport Me.crpConsulta, "rptRecibo", db.sconection, Array("FactPagadas", "Valores"), Array(Array("pClave", Me.grdComprobantes.TextMatrix(Me.grdComprobantes.row, 0)))
    
    Me.MousePointer = 0
    
    Me.cmdCRecibo.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub datDesde_Change()

    Call cargarComprobantes
    
End Sub

Private Sub datHasta_Change()

    Call cargarComprobantes
    
End Sub

Private Sub Form_Load()
Dim intAnchos As Variant
Dim strTitulos As Variant

    strTitulos = Array("Clave", "Comprobante", "Pref", "Nro", "Fecha", "Regist", "Importe", "Cancelado", "CmpID", "NroCompConta")
    intAnchos = Array(0, 1500, 500, 900, 1000, 1000, 850, 850, 0, 0)
    
    modGrid.makeGrid Me.grdComprobantes, strTitulos, intAnchos, 0, 1, flexSelectionByRow
    
    cliente.fillCombo Me.cboCliente, db
     
    Me.datDesde.value = Date
    Me.datHasta.value = Date
    
End Sub

Private Sub cboCliente_Click()

    Call cargarComprobantes

End Sub

