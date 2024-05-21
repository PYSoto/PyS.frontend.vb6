VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsVDDLAlterArtProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "vwalterartprov"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "vwalterartprov.clave"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = "vwalterartprov.codigo"

End Property

Public Property Get descripcion() As String

    descripcion = "descripcion"

End Property

Public Property Get descripcion_t() As String

    descripcion_t = "vwalterartprov.descripcion"

End Property

Public Property Get razon() As String

    razon = "razon"

End Property

Public Property Get razon_t() As String

    razon_t = "vwalterartprov.razon"

End Property

