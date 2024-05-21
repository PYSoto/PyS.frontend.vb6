VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLCompAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "compafip"

End Property

Public Property Get comprobanteID() As String

    comprobanteID = "caf_id"

End Property

Public Property Get comprobanteID_t() As String

    comprobanteID_t = "compafip.caf_id"

End Property

Public Property Get nombre() As String

    nombre = "caf_nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "compafip.caf_nombre"

End Property

Public Property Get etiqueta() As String

    etiqueta = "caf_label"

End Property

Public Property Get etiqueta_t() As String

    etiqueta_t = "compafip.caf_label"

End Property

