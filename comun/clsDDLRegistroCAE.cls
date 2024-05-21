VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLRegistroCAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "registrocae"

End Property

Public Property Get tcoID() As String

    tcoID = "rec_tco_id"

End Property

Public Property Get tcoID_t() As String

    tcoID_t = table & ".rec_tco_id"

End Property

Public Property Get prefijo() As String

    prefijo = "rec_prefijo"

End Property

Public Property Get prefijo_t() As String

    prefijo_t = table & ".rec_prefijo"

End Property

Public Property Get nroComprob() As String

    nroComprob = "rec_nrocomprob"

End Property

Public Property Get nroComprob_t() As String

    nroComprob_t = table & ".rec_nrocomprob"

End Property

Public Property Get cliID() As String

    cliID = "rec_cli_id"

End Property

Public Property Get cliID_t() As String

    cliID_t = table & ".rec_cli_id"

End Property

Public Property Get total() As String

    total = "rec_total"

End Property

Public Property Get total_t() As String

    total_t = table & ".rec_total"

End Property

Public Property Get exento() As String

    exento = "rec_exento"

End Property

Public Property Get exento_t() As String

    exento_t = table & ".rec_exento"

End Property

Public Property Get neto() As String

    neto = "rec_neto"

End Property

Public Property Get neto_t() As String

    neto_t = table & ".rec_neto"

End Property

Public Property Get neto105() As String

    neto105 = "rec_neto105"

End Property

Public Property Get neto105_t() As String

    neto105_t = table & ".rec_neto105"

End Property

Public Property Get iva() As String

    iva = "rec_iva"

End Property

Public Property Get iva_t() As String

    iva_t = table & ".rec_iva"

End Property

Public Property Get iva105() As String

    iva105 = "rec_iva105"

End Property

Public Property Get iva105_t() As String

    iva105_t = table & ".rec_iva105"

End Property

Public Property Get cae() As String

    cae = "rec_cae"

End Property

Public Property Get cae_t() As String

    cae_t = table & ".rec_cae"

End Property

Public Property Get fecha() As String

    fecha = "rec_fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = table & ".rec_fecha"

End Property

Public Property Get caevenc() As String

    caevenc = "rec_caevenc"

End Property

Public Property Get caevenc_t() As String

    caevenc_t = table & ".rec_caevenc"

End Property

Public Property Get barras() As String

    barras = "rec_barras"

End Property

Public Property Get barras_t() As String

    barras_t = table & ".rec_barras"

End Property

Public Property Get tipodocumento() As String

    tipodocumento = "tipo_documento"

End Property

Public Property Get tipodocumento_t() As String

    tipodocumento_t = table & ".tipo_documento"

End Property

Public Property Get numerodocumento() As String

    numerodocumento = "numero_documento"

End Property

Public Property Get numerodocumento_t() As String

    numerodocumento_t = table & ".numero_documento"

End Property

Public Property Get clientemovimientoIdasociado() As String

    clientemovimientoIdasociado = "cliente_movimiento_id_asociado"

End Property

Public Property Get clientemovimientoIdasociado_t() As String

    clientemovimientoIdasociado_t = table & ".cliente_movimiento_id_asociado"

End Property

Public Property Get recID() As String

    recID = "rec_id"

End Property

Public Property Get recID_t() As String

    recID_t = table & ".rec_id"

End Property


