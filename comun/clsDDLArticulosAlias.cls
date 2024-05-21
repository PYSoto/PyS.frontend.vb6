VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulosAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulosalias"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "articulosalias.clave"

End Property

Public Property Get artid() As String

    artid = "aal_art_id"

End Property

Public Property Get artid_t() As String

    artid_t = "articulosalias.aal_art_id"

End Property

Public Property Get alias() As String

    alias = "aal_alias"

End Property

Public Property Get alias_t() As String

    alias_t = "articulosalias.aal_alias"

End Property

Public Property Get prvid() As String

    prvid = "aal_prv_id"

End Property

Public Property Get prvid_t() As String

    prvid_t = "articulosalias.aal_prv_id"

End Property

Public Property Get preciocompra() As String

    preciocompra = "aal_preciocompra"

End Property

Public Property Get preciocompra_t() As String

    preciocompra_t = "articulosalias.aal_preciocompra"

End Property

