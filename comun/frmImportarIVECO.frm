VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarIVECO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Precios IVECO"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   13710
   Begin VB.TextBox txtCotizacion 
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cboCambios 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   3615
   End
   Begin ComctlLib.ProgressBar prbProgreso 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7320
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdIveco 
      Height          =   5775
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRevisar 
      Caption         =   "Revisar"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin ComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Width           =   13710
      _ExtentX        =   24183
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
   Begin MSComDlg.CommonDialog cdlArchivo 
      Left            =   10200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   ". . ."
      Height          =   255
      Left            =   12720
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   13215
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11760
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cotización"
      Height          =   195
      Left            =   7920
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Filtro"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   330
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Archivo"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmImportarIVECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colArtFile As Collection
Private colArtDB As Collection

Private objArtFile As New clsDAOArticulos
Private objArtDB As New clsDAOArticulos

Private Sub loadFileTxt()
Dim strLinea As String

    Set colArtFile = New Collection
    
    Open Me.txtArchivo.Text For Input As #1
    Line Input #1, strLinea
    Do While Not EOF(1)
        Set objArtFile = parseArticulo(strLinea, Val(Me.txtCotizacion.Text))
        If Not modCollection.collectionExistElement(colArtFile, "k." & objArtFile.codigo) Then colArtFile.add objArtFile, "k." & objArtFile.codigo
        Line Input #1, strLinea
    Loop
    Close #1

End Sub

Private Sub loadFileXls()
Dim xlsSV As Object
Dim xlsWB As Object
Dim xlsAS As Object

Dim row As Long
Dim rowcount As Long

Dim intCol As Integer
Dim intColCount As Integer
Dim intCatalogo As Integer
Dim intOrigen As Integer
Dim intSinIva As Integer
Dim intDescuento As Integer
Dim intDenominacion As Integer
Dim intFecha As Integer
Dim intModelo As Integer

Dim curCotizacion As Currency

On Error Resume Next

    Me.cmdImportar.Enabled = False
    Me.cmdRevisar.Enabled = False
    Me.cmdSalir.Enabled = False
    Me.cmdArchivo.Enabled = False
    
    Set colArtFile = New Collection
    
    Set xlsSV = CreateObject("Excel.Application")
    Set xlsWB = xlsSV.Workbooks.Open(Me.txtArchivo.Text)
    xlsSV.Visible = False
    Set xlsAS = xlsWB.Worksheets(1)
        
    rowcount = xlsAS.UsedRange.Rows.Count
    intColCount = xlsAS.UsedRange.Columns.Count
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Max = rowcount + 2
    Me.prbProgreso.Value = 1
    
    For intCol = 1 To intColCount
        Debug.Print xlsAS.cells(1, intCol).Value
        Select Case xlsAS.cells(1, intCol).Value
            Case "Catálogo"
                intCatalogo = intCol
            Case "O"
                intOrigen = intCol
            Case "s/IVA"
                intSinIva = intCol
            Case "D"
                intDescuento = intCol
            Case "Denominación"
                intDenominacion = intCol
            Case "Fecha"
                intFecha = intCol
            Case "Modelo"
                intModelo = intCol
        End Select
    Next intCol
    
    curCotizacion = Val(Me.txtCotizacion.Text)
    For row = 2 To rowcount
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        DoEvents
        Set objArtFile = New clsDAOArticulos
        With objArtFile
            .codigo = Trim(xlsAS.cells(row, intCatalogo).Value) & ".I"
            .descripcion = Trim(modDB.cleanGarbage(xlsAS.cells(row, intDenominacion).Value))
            .preciolistasiniva = a2Decimales(Val(xlsAS.cells(row, intSinIva).Value) * curCotizacion)
            .origen = xlsAS.cells(row, intOrigen).Value
            .descuento = Trim(xlsAS.cells(row, intDescuento).Value)
            .fechaactualizacion = CDate(xlsAS.cells(row, intFecha).Value)
            .modelocamion = Trim(xlsAS.cells(row, intModelo).Value)
            .preciolistasinivausd = Val(xlsAS.cells(row, intSinIva).Value)
        End With
        If Not modCollection.collectionExistElement(colArtFile, "k." & objArtFile.codigo) Then colArtFile.add objArtFile, "k." & objArtFile.codigo
    Next row

    Me.cmdImportar.Enabled = True
    Me.cmdRevisar.Enabled = True
    Me.cmdSalir.Enabled = True
    Me.cmdArchivo.Enabled = True

End Sub

Private Sub listAll()
Dim blnAdd As Boolean

On Error Resume Next

    Me.prbProgreso.Min = 1
    Me.prbProgreso.Value = 1
    Me.prbProgreso.Max = colArtFile.Count + 1

    Me.grdIveco.Rows = 1
    Me.grdIveco.Redraw = False
    For Each objArtFile In colArtFile
        Set objArtDB = New clsDAOArticulos
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        
        blnAdd = False
        
        If modCollection.collectionExistElement(colArtDB, "k." & objArtFile.codigo) Then Set objArtDB = colArtDB("k." & objArtFile.codigo)

        With objArtDB
            If objArtFile.descripcion <> .descripcion Then blnAdd = True
            If objArtFile.preciolistasiniva <> .preciolistasiniva Then blnAdd = True
            If objArtFile.origen <> .origen Then blnAdd = True
            If objArtFile.descuento <> .descuento Then blnAdd = True

            If blnAdd Then
                Me.grdIveco.AddItem modGrid.array2itemGrid(Array(objArtFile.codigo, objArtFile.descripcion, Format(objArtFile.preciolistasiniva, "0.00"), objArtFile.origen, objArtFile.descuento, objArtFile.fechaactualizacion, "", .descripcion, Format(.preciolistasiniva, "0.00"), .origen, .descuento, .fechaactualizacion))
                modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 6, modColor.colorAzul
                If objArtFile.descripcion <> .descripcion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 7, modColor.colorRojo
                If objArtFile.preciolistasiniva <> .preciolistasiniva Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 8, modColor.colorRojo
                If objArtFile.origen <> .origen Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 9, modColor.colorRojo
                If objArtFile.descuento <> .descuento Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 10, modColor.colorRojo
                If objArtFile.fechaactualizacion <> .fechaactualizacion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 11, modColor.colorRojo
            End If
        End With
    Next
    Me.grdIveco.Redraw = True

End Sub

Private Sub listUp()
Dim blnAdd As Boolean

    Me.prbProgreso.Min = 1
    Me.prbProgreso.Value = 1
    Me.prbProgreso.Max = colArtFile.Count + 1

    Me.grdIveco.Rows = 1
    Me.grdIveco.Redraw = False
    For Each objArtFile In colArtFile
        Set objArtDB = New clsDAOArticulos
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        
        blnAdd = False
        
        If modCollection.collectionExistElement(colArtDB, "k." & objArtFile.codigo) Then Set objArtDB = colArtDB("k." & objArtFile.codigo)

        With objArtDB
            If objArtFile.preciolistasiniva > .preciolistasiniva Then blnAdd = True

            If blnAdd Then
                Me.grdIveco.AddItem modGrid.array2itemGrid(Array(objArtFile.codigo, objArtFile.descripcion, Format(objArtFile.preciolistasiniva, "0.00"), objArtFile.origen, objArtFile.descuento, objArtFile.fechaactualizacion, "", .descripcion, Format(.preciolistasiniva, "0.00"), .origen, .descuento, .fechaactualizacion))
                modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 6, modColor.colorAzul
                If objArtFile.descripcion <> .descripcion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 7, modColor.colorRojo
                If objArtFile.preciolistasiniva <> .preciolistasiniva Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 8, modColor.colorRojo
                If objArtFile.origen <> .origen Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 9, modColor.colorRojo
                If objArtFile.descuento <> .descuento Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 10, modColor.colorRojo
                If objArtFile.fechaactualizacion <> .fechaactualizacion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 11, modColor.colorRojo
            End If
        End With
    Next
    Me.grdIveco.Redraw = True

End Sub

Private Sub listDown()
Dim blnAdd As Boolean

    Me.prbProgreso.Min = 1
    Me.prbProgreso.Value = 1
    Me.prbProgreso.Max = colArtFile.Count + 1

    Me.grdIveco.Rows = 1
    Me.grdIveco.Redraw = False
    For Each objArtFile In colArtFile
        Set objArtDB = New clsDAOArticulos
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        
        blnAdd = False
        
        If modCollection.collectionExistElement(colArtDB, "k." & objArtFile.codigo) Then Set objArtDB = colArtDB("k." & objArtFile.codigo)

        With objArtDB
            If objArtFile.preciolistasiniva < .preciolistasiniva Then blnAdd = True

            If blnAdd Then
                Me.grdIveco.AddItem modGrid.array2itemGrid(Array(objArtFile.codigo, objArtFile.descripcion, Format(objArtFile.preciolistasiniva, "0.00"), objArtFile.origen, objArtFile.descuento, objArtFile.fechaactualizacion, "", .descripcion, Format(.preciolistasiniva, "0.00"), .origen, .descuento, .fechaactualizacion))
                modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 6, modColor.colorAzul
                If objArtFile.descripcion <> .descripcion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 7, modColor.colorRojo
                If objArtFile.preciolistasiniva <> .preciolistasiniva Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 8, modColor.colorRojo
                If objArtFile.origen <> .origen Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 9, modColor.colorRojo
                If objArtFile.descuento <> .descuento Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 10, modColor.colorRojo
                If objArtFile.fechaactualizacion <> .fechaactualizacion Then modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 11, modColor.colorRojo
            End If
        End With
    Next
    Me.grdIveco.Redraw = True

End Sub

Private Sub listRemove()
Dim blnAdd As Boolean

Dim colImp As Collection

Dim objImp As New clsDAOImportaciones

Dim objArI As New clsDAOArticulosImport

    objImp.findLast
    
    Set colImp = objArI.collectionByFecha(objImp.fecha)

    Me.prbProgreso.Min = 1
    Me.prbProgreso.Value = 1
    Me.prbProgreso.Max = colImp.Count + 1

    Me.grdIveco.Rows = 1
    Me.grdIveco.Redraw = False
    For Each objArI In colImp
        Set objArtFile = New clsDAOArticulos
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
        
        blnAdd = True
        
        Set objArtFile = Nothing
        
        If modCollection.collectionExistElement(colArtFile, "k." & objArI.articuloID) Then
            Set objArtFile = colArtFile("k." & objArI.articuloID)
            blnAdd = False
        End If

        With objArI
            If blnAdd Then
                Me.grdIveco.AddItem modGrid.array2itemGrid(Array(objArI.articuloID, objArtFile.descripcion, Format(objArtFile.preciolistasiniva, "0.00"), objArtFile.origen, objArtFile.descuento, objArtFile.fechaactualizacion, "", .descripcion, Format(.preciolistasiniva, "0.00"), .origen, .descuento, .fechaactualizacion))
                modGrid.setColourGrid Me.grdIveco, Me.grdIveco.Rows - 1, 6, modColor.colorAzul
            End If
        End With
    Next
    Me.grdIveco.Redraw = True

End Sub

Private Function parseArticulo(pLinea As String, pCotizacion As Currency) As clsDAOArticulos
Dim objArt As New clsDAOArticulos

'    With objArt
'        .codigo = Trim(Left(pLinea, 15))
'        If Left(.codigo, 1) = "*" Then .codigo = Trim(Mid(.codigo, 2))
'        .codigo = Trim(.codigo)
'        If Mid(pLinea, 16, 1) <> "N" And Mid(pLinea, 16, 1) <> "0" Then .codigo = .codigo & "." & Mid(pLinea, 16, 1)
'        .descripcion = Trim(Mid(pLinea, 42, 24))
'        .preciolistasiniva = Val(Mid(pLinea, 18, 10)) / 100 * pCotizacion
'        .origen = Mid(pLinea, 16, 1)
'        .descuento = Trim(Mid(pLinea, 40, 1))
'        .fechaactualizacion = CDate(Mid(pLinea, 75, 2) & "/" & Mid(pLinea, 77, 2) & "/" & Mid(pLinea, 79, 4))
'        .modelocamion = Trim(Mid(pLinea, 88))
'    End With
        
    With objArt
        .codigo = Trim(Left(pLinea, 15))
        If Left(.codigo, 1) = "*" Then .codigo = Trim(Mid(.codigo, 2))
        .codigo = Trim(.codigo) & ".I"
        .descripcion = Trim(Mid(pLinea, 42, 24))
        .preciolistasiniva = a2Decimales(Val(Mid(pLinea, 18, 10)) / 100 * pCotizacion)
        .origen = Mid(pLinea, 16, 1)
        .descuento = Trim(Mid(pLinea, 40, 1))
        '.fechaactualizacion = Date
        .fechaactualizacion = CDate(Mid(pLinea, 67, 2) & "/" & Mid(pLinea, 69, 2) & "/" & Mid(pLinea, 71, 4))
        .modelocamion = Trim(Mid(pLinea, 88))
        .preciolistasinivausd = Val(Mid(pLinea, 18, 10)) / 100
    End With
            
    Set parseArticulo = objArt

End Function

Private Sub recalcularPrecios()
Dim curDescuento As Currency

    With objArtDB
        .preciolistaconiva = a2Decimales(.preciolistasiniva * 1.21)
        .precioventaconiva = a2Decimales(.preciolistaconiva)
        .precioventasiniva = a2Decimales(.preciolistasiniva)

        ' Determina precio de compra
        curDescuento = 0.25

        Select Case .descuento
            Case "A", "9"
                curDescuento = 0.05
            Case "S", "7"
                curDescuento = 0.1
            Case "C", "6"
                curDescuento = 0.12
        End Select
    
        .preciocomprasiniva = a2Decimales(.precioventasiniva * (1 - curDescuento))
        
        ' Determina precio de venta
        curDescuento = 0.1
        
        Select Case .descuento
            Case "A", "9"
                curDescuento = 0
            Case "S", "7"
                curDescuento = 0
            Case "C", "6"
                curDescuento = 0
        End Select
        
        .precioventaconiva = a2Decimales(.preciolistaconiva * (1 - curDescuento))
        .precioventasiniva = a2Decimales(.preciolistasiniva * (1 - curDescuento))
    End With
    
End Sub

Private Sub cmdArchivo_Click()

    Set colArtFile = Nothing

    With cdlArchivo
        .Filter = "Lista|*.xls"
        .ShowOpen
        Me.txtArchivo.Text = .FileName
    End With
    
    If Trim(Me.txtArchivo.Text) = "" Then Exit Sub
    
    Me.MousePointer = 11
    
    Me.stbEstado.SimpleText = "Cargando Archivo ..."
    
    loadFileXls
    
    Me.stbEstado.SimpleText = "Archivo LOADED ..."
    
    Me.MousePointer = 0

End Sub

Private Sub cmdImportar_Click()
Dim objImp As New clsDAOImportaciones

Dim objArI As New clsDAOArticulosImport

Dim objCot As New clsDAOCotizacion

On Error Resume Next

    Me.txtCotizacion.Text = Format(Val(Me.txtCotizacion.Text), "0.00")
    
    If Val(Me.txtCotizacion.Text) = 0 Then
        MsgBox "ERROR: SIN Cotización"
        Exit Sub
    End If
    
    If Trim(Me.txtArchivo.Text) = "" Then Exit Sub
    
    Me.cmdImportar.Enabled = False
    Me.cmdRevisar.Enabled = False
    Me.cmdSalir.Enabled = False
    Me.cmdArchivo.Enabled = False
    
    Me.MousePointer = 11
    
    Me.stbEstado.SimpleText = "Importando ..."
    
    With objImp
        .fecha = Date
        .save
    End With
    
    With objCot
        .fecha = Date
        
        .findByFecha
        
        .fecha = Date
        .usdventa = Val(Me.txtCotizacion.Text)
        
        .save
    End With
    
    gArtMirror.saveCotizacion objCot, db
    
    Set colArtDB = Nothing
    Set colArtDB = objArtFile.collectionAll
    
    Me.prbProgreso.Min = 1
    Me.prbProgreso.Value = 1
    Me.prbProgreso.Max = colArtFile.Count + 1
    
    For Each objArtFile In colArtFile
        DoEvents
        Me.prbProgreso.Value = Me.prbProgreso.Value + 1
    
        Set objArtDB = Nothing
        Set objArtDB = New clsDAOArticulos
        objArtDB.codigo = objArtFile.codigo
        
        If modCollection.collectionExistElement(colArtDB, "k." & objArtFile.codigo) Then Set objArtDB = colArtDB("k." & objArtFile.codigo)
        
        With objArtDB
            .descripcion = objArtFile.descripcion
            .preciolistasiniva = objArtFile.preciolistasiniva
            .origen = objArtFile.origen
            .descuento = objArtFile.descuento
            .modelocamion = objArtFile.modelocamion
            .fechaactualizacion = objArtFile.fechaactualizacion
            .preciolistasinivausd = objArtFile.preciolistasinivausd
            .cotizacionID = objCot.cotizacionID
            
            Call recalcularPrecios
            
            If .clave > 0 Then .update Else .add
            
            gArtMirror.saveArticulo objArtDB, db
            
            Me.stbEstado.SimpleText = "Actualizando: (" & .codigo & ") " & .descripcion
        End With
        
        With objArI
            .fecha = Date
            .articuloID = objArtFile.codigo
            .descripcion = objArtFile.descripcion
            .preciolistasiniva = objArtFile.preciolistasiniva
            .origen = objArtFile.origen
            .descuento = objArtFile.descuento
            .fechaactualizacion = objArtFile.fechaactualizacion
            .cotizacionID = objCot.cotizacionID
            
            .save
        End With
    Next
    
    Me.MousePointer = 0
    
    Me.cmdImportar.Enabled = True
    Me.cmdRevisar.Enabled = True
    Me.cmdSalir.Enabled = True
    Me.cmdArchivo.Enabled = True
    
    Me.stbEstado.SimpleText = "Proceso TERMINADO"
    
End Sub

Private Sub cmdRevisar_Click()

    Me.txtCotizacion.Text = Format(Val(Me.txtCotizacion.Text), "0.00")
    
    If Trim(Me.txtArchivo.Text) = "" Then Exit Sub
    
    If Me.cboCambios.ListIndex < 0 Then Exit Sub
    
    If Val(Me.txtCotizacion.Text) = 0 Then
        MsgBox "ERROR: SIN Cotización"
        Exit Sub
    End If
    
    Me.cmdImportar.Enabled = False
    Me.cmdRevisar.Enabled = False
    Me.cmdSalir.Enabled = False
    Me.cmdArchivo.Enabled = False
    
    Me.MousePointer = 11
    
    Me.stbEstado.SimpleText = "Revisando..."
    
    Set colArtDB = objArtFile.collectionAll
    
    Select Case Me.cboCambios.ListIndex
        Case 0:
            listAll
        Case 1:
            listUp
        Case 2:
            listDown
        Case 3:
            listRemove
    End Select

    Me.MousePointer = 0
    
    Me.stbEstado.SimpleText = ""
    
    Me.cmdImportar.Enabled = True
    Me.cmdRevisar.Enabled = True
    Me.cmdSalir.Enabled = True
    Me.cmdArchivo.Enabled = True

End Sub

Private Sub cmdSalir_Click()

    Set colArtDB = Nothing
    Set colArtFile = Nothing
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    modGrid.makeGrid2 Me.grdIveco, Array(Array("Código", 1500), Array("Descripcion", 3000), Array("Pr s/IVA", 1000), Array("O", 250), Array("D", 250), Array("Fecha", 1000), Array("", 100), Array("Descripcion", 3000), Array("Pr s/IVA", 1000), Array("O", 250), Array("D", 250), Array("Fecha", 1000)), 0, 1, flexSelectionByRow
    
    Me.cboCambios.AddItem "Todos los cambios"
    Me.cboCambios.AddItem "Aumento de Precios"
    Me.cboCambios.AddItem "Descenso de Precios"
    Me.cboCambios.AddItem "No Informados"
    
    Me.cboCambios.ListIndex = 0
    
End Sub

Private Sub txtCotizacion_GotFocus()

    marcarseleccion Me.txtCotizacion
    
End Sub

Private Sub txtCotizacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCotizacion_LostFocus
    End If
    
End Sub

Private Sub txtCotizacion_LostFocus()

    If Trim(Me.txtArchivo.Text) = "" Then Exit Sub
    
    Me.MousePointer = 11
    
    loadFileXls
    
    Me.stbEstado.SimpleText = "Archivo LOADED ..."
    
    Me.MousePointer = 0
    
End Sub
