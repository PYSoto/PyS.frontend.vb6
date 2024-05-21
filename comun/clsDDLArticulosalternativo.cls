VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulosAlter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulosalternativo"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "articulosalternativo.clave"

End Property

Public Property Get artid() As String

    artid = "ara_art_id"

End Property

Public Property Get artid_t() As String

    artid_t = "articulosalternativo.ara_art_id"

End Property

Public Property Get artidalternativo() As String

    artidalternativo = "ara_art_id_alternativo"

End Property

Public Property Get artidalternativo_t() As String

    artidalternativo_t = "articulosalternativo.ara_art_id_alternativo"

End Property

