VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function saveCompra(proveedormov As clsDAOProveedorMov, pDB As clsDB) As Boolean

    saveCompra = False
    pDB.beginTrans
    
    If Not proveedormov.add(pDB) Then GoTo rollback
    
    pDB.commitTrans
    saveCompra = True
    Exit Function
    
rollback:
    pDB.rollback

End Function

Public Function savePago(proveedormov As clsDAOProveedorMov, proveedor As clsDAOProveedor, proveedormovs As Collection, valormovs As Collection, pDB As clsDB) As Boolean
Dim valormov As clsDAOValorMov
Dim proveedormovloc As clsDAOProveedorMov
Dim proveedorpago As clsDAOProveedorPago

    savePago = False
    pDB.beginTrans
    If Not proveedormov.add(pDB) Then GoTo rollback
    For Each valormov In valormovs
        valormov.proveedormovimientoID = proveedormov.proveedormovimientoID
        valormov.importe = -valormov.importe
        
        If Not valormov.add(pDB) Then GoTo rollback
    Next
    If Not proveedormovs Is Nothing Then
        For Each proveedormovloc In proveedormovs
            If proveedormovloc.pago <> 0 Then
                Set proveedorpago = New clsDAOProveedorPago
                proveedorpago.proveedorMovimientoIDDeuda = proveedormovloc.proveedormovimientoID
                proveedorpago.proveedorMovimientoIDAplicado = proveedormov.proveedormovimientoID
                proveedorpago.importeAplicado = proveedormovloc.pago
                
                If Not proveedorpago.save(pDB) Then GoTo rollback
                
                proveedormovloc.totalCancelado = proveedormovloc.totalCancelado + proveedormovloc.pago
                If Not proveedormovloc.save(pDB) Then GoTo rollback
                
                proveedormov.totalCancelado = proveedormov.totalCancelado - proveedormovloc.pago
            End If
        Next
        If Not proveedormov.save(pDB) Then GoTo rollback
    End If
    
    pDB.commitTrans
    savePago = True
    Exit Function
    
rollback:
    pDB.rollback
    
End Function

Public Function updateSaldoCCTope(proveedorID As Long, fechatope As Date, pDB As clsDB) As Boolean
Dim proveedormov As New clsDAOProveedorMov
Dim proveedorsaldo As New clsDAOProveedorSaldo

    proveedorsaldo.proveedorID = proveedorID
    proveedorsaldo.fecha = fechatope
    proveedorsaldo.saldo = 0
    
    For Each proveedormov In proveedormov.collectionCtaCteByProveedorID(proveedorID, pDB, fechatope)
        proveedorsaldo.saldo = proveedorsaldo.saldo + proveedormov.total
    Next
    
    updateSaldoCCTope = proveedorsaldo.save(pDB)
    
End Function

Public Function nextOP(pDB As clsDB) As Long
Dim proveedormov As New clsDAOProveedorMov
    
    proveedormov.findLastOPago pDB
    
    nextOP = 1 + proveedormov.nroComprobante

End Function

Public Function nextComprobante(comprobanteID As Integer, pDB As clsDB) As Long
Dim proveedormov As New clsDAOProveedorMov
    
    proveedormov.comprobanteID = comprobanteID
    proveedormov.findLastByComprobante pDB
    
    nextComprobante = 1 + proveedormov.nroComprobante

End Function

Public Function deleteComprobante(proveedormov As clsDAOProveedorMov, pDB As clsDB) As Boolean
Dim proveedorpago As New clsDAOProveedorPago
Dim proveedormovloc As New clsDAOProveedorMov
Dim valormov As New clsDAOValorMov

    deleteComprobante = False

    pDB.beginTrans

    ' Eliminar aplicacion
    For Each proveedorpago In proveedorpago.collectionByPago(proveedormov.proveedormovimientoID, pDB)
        proveedormovloc.proveedormovimientoID = proveedorpago.proveedorMovimientoIDDeuda
        proveedormovloc.findByProveedorMovimientoID pDB
        ' Resto el aplicado al comprobante adeudado
        proveedormovloc.totalCancelado = proveedormovloc.totalCancelado - proveedorpago.importeAplicado
        If Not proveedormovloc.update(pDB) Then GoTo rollback
        ' Elimina el registro de asociacion
        If Not proveedorpago.delete(pDB) Then GoTo rollback
    Next
    ' Eliminar valores
    For Each valormov In valormov.collectionByProveedorMovimientoID(proveedormov.proveedormovimientoID, pDB)
        If Not valormov.delete(pDB) Then GoTo rollback
    Next
    ' Eliminar comprobante
    If Not proveedormov.delete(pDB) Then GoTo rollback
    
    pDB.commitTrans
    
    deleteComprobante = True
    
    Exit Function
    
rollback:
    pDB.rollback
    
End Function


