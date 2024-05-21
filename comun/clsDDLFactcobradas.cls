VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLFactcobradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "factcobradas"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "factcobradas.clave"

End Property

Public Property Get clavemovc() As String

    clavemovc = "clavemovc"

End Property

Public Property Get clavemovc_t() As String

    clavemovc_t = "factcobradas.clavemovc"

End Property

Public Property Get clavepago() As String

    clavepago = "clavepago"

End Property

Public Property Get clavepago_t() As String

    clavepago_t = "factcobradas.clavepago"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "factcobradas.importe"

End Property

