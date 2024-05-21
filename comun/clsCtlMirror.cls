VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vDB As New clsDB

Private Function existMirror(pDB As clsDB) As Boolean
Dim objArM As New clsDAOArticulosMirror
    
    objArM.findLast pDB
    
    If objArM.id = 0 Then
        existMirror = False
        Exit Function
    End If
    
    With vDB
        .dsn = objArM.dsn
        .uid = objArM.uid
        .pwd = objArM.pwd
        .ip = objArM.ip
        .database = objArM.database
        
        .openDB
    End With
    
    existMirror = True
    
End Function

Public Function saveArticulo(pArt As clsDAOArticulos, pDB As clsDB) As Boolean
Dim objArt As clsDAOArticulos

    saveArticulo = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objArt = pArt.clone
    
    saveArticulo = objArt.save(vDB)
    
End Function

Public Function saveCotizacion(pCot As clsDAOCotizacion, pDB As clsDB) As Boolean
Dim objCot As clsDAOCotizacion

    saveCotizacion = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objCot = pCot.clone
    
    saveCotizacion = objCot.save(vDB)
    
End Function

Public Function deleteArticulo(pArt As clsDAOArticulos, pDB As clsDB) As Boolean
Dim objArt As clsDAOArticulos

    deleteArticulo = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objArt = pArt.clone
    
    deleteArticulo = objArt.delete(vDB)
    
End Function

Public Function saveArticuloUbic(pArtUbic As clsDAOArticulosUbic, pDB As clsDB) As Boolean
Dim objAU As clsDAOArticulosUbic

    saveArticuloUbic = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objAU = pArtUbic.clone
    
    saveArticuloUbic = objAU.save(vDB)
    
End Function

Public Function saveArticuloAlias(pArtAlias As clsDAOArticulosAlias, pDB As clsDB) As Boolean
Dim objAA As clsDAOArticulosAlias

    saveArticuloAlias = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objAA = pArtAlias.clone
    
    With objAA
        .findByAlias objAA.alias, vDB
        
        .artid = pArtAlias.artid
        .prvid = pArtAlias.prvid
        .preciocompra = pArtAlias.preciocompra
     
        saveArticuloAlias = .save(vDB)
    End With
    
End Function

Public Function deleteArticuloAlias(pArtAlias As clsDAOArticulosAlias, pDB As clsDB) As Boolean
Dim objAA As clsDAOArticulosAlias

    deleteArticuloAlias = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objAA = pArtAlias.clone
    
    objAA.findByAlias objAA.alias, vDB
    
    deleteArticuloAlias = objAA.delete(vDB)
    
End Function

Public Function saveArticuloAlter(pArtAlter As clsDAOArticulosAlter, pDB As clsDB) As Boolean
Dim objAA As clsDAOArticulosAlter

    saveArticuloAlter = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objAA = pArtAlter.clone
    
    saveArticuloAlter = objAA.save(vDB)
    
End Function

Public Function deleteArticuloAlter(pArtAlter As clsDAOArticulosAlter, pDB As clsDB) As Boolean
Dim objAA As clsDAOArticulosAlter

    deleteArticuloAlter = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set objAA = pArtAlter.clone
    
    objAA.findByUnique objAA.artid, objAA.artidalternativo, vDB
    
    deleteArticuloAlter = objAA.delete(vDB)
    
End Function

Public Function saveProveedor(proveedor As clsDAOProveedor, pDB As clsDB) As Boolean
Dim proveedorlocal As clsDAOProveedor

    saveProveedor = True
    
    If Not existMirror(pDB) Then Exit Function
    
    Set proveedorlocal = proveedor.clone
    
    saveProveedor = proveedorlocal.save(vDB)
    
End Function

Public Sub depuraCliente(pCli As clsDAOClientes, pDB As clsDB)
Dim objCliR As New clsDAOClientes

Dim objMC As New clsDAOMovclie

Dim intCantL As Integer
Dim intCantR As Integer

    If Not existMirror(pDB) Then Exit Sub
    
    With objCliR
        .findByCUIT pCli.cuit, vDB
        
        If .clave = 0 Then
            pCli.facturable = 1
            pCli.save pDB
            Exit Sub
        End If
    End With
    
    intCantL = objMC.collectionByCliID(pCli.codigo, pDB).Count
    intCantR = objMC.collectionByCliID(objCliR.codigo, vDB).Count
    
    If intCantL >= intCantR Then
        pCli.facturable = 1
        pCli.save pDB
        
        objCliR.facturable = 0
        objCliR.save vDB
        Exit Sub
    End If
    
    pCli.facturable = 0
    pCli.save pDB
    
    objCliR.facturable = 1
    objCliR.save vDB
    
End Sub
