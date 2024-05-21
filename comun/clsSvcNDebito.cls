VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSvcNDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub autoCorrectNDebito(negocioID As Integer, tipoComprobanteID As Integer, prefijo As Integer, nroComprob As Long, db As clsDB)
Dim objMCl As New clsDAOMovclie

Dim neto As Currency
Dim exento As Currency

    If tipoComprobanteID < 40 Or tipoComprobanteID > 41 Then Exit Sub
    
    objMCl.findByComprobante negocioID, tipoComprobanteID, prefijo, nroComprob, db
    
    If objMCl.montoexento <> 0 Then Exit Sub
    
    neto = objMCl.montoiva / 0.21
    exento = objMCl.importe - objMCl.montoiva - neto
    objMCl.neto = neto
    objMCl.netosindescuento = neto
    objMCl.montoexento = exento
    
    objMCl.save db

End Sub
