VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulosUbic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulosubic"

End Property

Public Property Get artid() As String

    artid = "aub_art_id"

End Property

Public Property Get artid_t() As String

    artid_t = "articulosubic.aub_art_id"

End Property

Public Property Get ubicacion() As String

    ubicacion = "aub_ubicacion"

End Property

Public Property Get ubicacion_t() As String

    ubicacion_t = "articulosubic.aub_ubicacion"

End Property

Public Property Get id() As String

    id = "aub_id"

End Property

Public Property Get id_t() As String

    id_t = "articulosubic.aub_id"

End Property

