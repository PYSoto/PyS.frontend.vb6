VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulosImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulosimport"

End Property

Public Property Get fecha() As String

    fecha = "art_fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "articulosimport.art_fecha"

End Property

Public Property Get articuloID() As String

    articuloID = "art_codigo"

End Property

Public Property Get articuloID_t() As String

    articuloID_t = "articulosimport.art_codigo"

End Property

Public Property Get descripcion() As String

    descripcion = "art_descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = "articulosimport.art_descripcion"

End Property

Public Property Get preciolistasiniva() As String

    preciolistasiniva = "art_preciolistasiniva"

End Property

Public Property Get preciolistasiniva_t() As String

    preciolistasiniva_t = "articulosimport.art_preciolistasiniva"

End Property

Public Property Get origen() As String

    origen = "art_origen"

End Property

Public Property Get origen_t() As String

    origen_t = "articulosimport.art_origen"

End Property

Public Property Get descuento() As String

    descuento = "art_descuento"

End Property

Public Property Get descuento_t() As String

    descuento_t = "articulosimport.art_descuento"

End Property

Public Property Get fechaactualizacion() As String

    fechaactualizacion = "art_fechaactualizacion"

End Property

Public Property Get fechaactualizacion_t() As String

    fechaactualizacion_t = "articulosimport.art_fechaactualizacion"

End Property

Public Property Get cotizacionID() As String

    cotizacionID = "cotizacion_id"

End Property

Public Property Get cotizacionID_t() As String

    cotizacionID_t = "articulosimport.cotizacion_id"

End Property

Public Property Get valorUSD() As String

    valorUSD = "valor_usd"

End Property

Public Property Get valorUSD_t() As String

    valorUSD_t = "articulosimport.valor_usd"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "articulosimport.clave"

End Property

