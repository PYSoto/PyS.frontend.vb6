VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "cotizacion"

End Property

Public Property Get cotizacionID() As String

    cotizacionID = "cotizacion_id"

End Property

Public Property Get cotizacionID_t() As String

    cotizacionID_t = "cotizacion.cotizacion_id"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "cotizacion.fecha"

End Property

Public Property Get usdcompra() As String

    usdcompra = "usd_compra"

End Property

Public Property Get usdcompra_t() As String

    usdcompra_t = "cotizacion.usd_compra"

End Property

Public Property Get usdventa() As String

    usdventa = "usd_venta"

End Property

Public Property Get usdventa_t() As String

    usdventa_t = "cotizacion.usd_venta"

End Property

