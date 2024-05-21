VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulos"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "articulos.clave"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = "articulos.codigo"

End Property

Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = "articulos.descripcion"

End Property

Public Property Get precioventaconiva() As String

    precioventaconiva = "art_precioventaconiva"

End Property

Public Property Get precioventaconiva_t() As String

    precioventaconiva_t = "articulos.art_precioventaconiva"

End Property

Public Property Get precioventasiniva() As String

    precioventasiniva = "art_precioventasiniva"

End Property

Public Property Get precioventasiniva_t() As String

    precioventasiniva_t = "articulos.art_precioventasiniva"

End Property

Public Property Get preciolistaconiva() As String

    preciolistaconiva = "art_preciolistaconiva"

End Property

Public Property Get preciolistaconiva_t() As String

    preciolistaconiva_t = "articulos.art_preciolistaconiva"

End Property

Public Property Get preciolistasiniva() As String

    preciolistasiniva = "art_preciolistasiniva"

End Property

Public Property Get preciolistasiniva_t() As String

    preciolistasiniva_t = "articulos.art_preciolistasiniva"

End Property

Public Property Get preciocomprasiniva() As String

    preciocomprasiniva = "art_preciocomprasiniva"

End Property

Public Property Get preciocomprasiniva_t() As String

    preciocomprasiniva_t = "articulos.art_preciocomprasiniva"

End Property

Public Property Get preciocomprasinivaanterior() As String

    preciocomprasinivaanterior = "art_preciocomprasinivaanterior"

End Property

Public Property Get preciocomprasinivaanterior_t() As String

    preciocomprasinivaanterior_t = "articulos.art_preciocomprasinivaanterior"

End Property

Public Property Get iva105() As String

    iva105 = "iva105"

End Property

Public Property Get iva105_t() As String

    iva105_t = "articulos.iva105"

End Property

Public Property Get exento() As String

    exento = "exento"

End Property

Public Property Get exento_t() As String

    exento_t = "articulos.exento"

End Property

Public Property Get modelocamion() As String

    modelocamion = "art_modelocamion"

End Property

Public Property Get modelocamion_t() As String

    modelocamion_t = "articulos.art_modelocamion"

End Property

Public Property Get fechaactualizacion() As String

    fechaactualizacion = "art_fechaactualizacion"

End Property

Public Property Get fechaactualizacion_t() As String

    fechaactualizacion_t = "articulos.art_fechaactualizacion"

End Property

Public Property Get origen() As String

    origen = "art_origen"

End Property

Public Property Get origen_t() As String

    origen_t = "articulos.art_origen"

End Property

Public Property Get descuento() As String

    descuento = "art_descuento"

End Property

Public Property Get descuento_t() As String

    descuento_t = "articulos.art_descuento"

End Property

Public Property Get prvid() As String

    prvid = "art_prv_id"

End Property

Public Property Get prvid_t() As String

    prvid_t = "articulos.art_prv_id"

End Property

Public Property Get ultimacompra() As String

    ultimacompra = "art_ultimacompra"

End Property

Public Property Get ultimacompra_t() As String

    ultimacompra_t = "articulos.art_ultimacompra"

End Property

Public Property Get marca() As String

    marca = "art_marca"

End Property

Public Property Get marca_t() As String

    marca_t = "articulos.art_marca"

End Property

Public Property Get catalogo() As String

    catalogo = "art_catalogo"

End Property

Public Property Get catalogo_t() As String

    catalogo_t = "articulos.art_catalogo"

End Property

Public Property Get preciolistasinivausd() As String

    preciolistasinivausd = "preciolistasinivausd"

End Property

Public Property Get preciolistasinivausd_t() As String

    preciolistasinivausd_t = "articulos.preciolistasinivausd"

End Property

Public Property Get cotizacionID() As String

    cotizacionID = "cotizacionID"

End Property

Public Property Get cotizacionID_t() As String

    cotizacionID_t = "articulos.cotizacionID"

End Property


