VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "feriados"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "feriados.clave"

End Property

Public Property Get fecha() As String

    fecha = "fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "feriados.fecha"

End Property

Public Property Get concepto() As String

    concepto = "concepto"

End Property

Public Property Get concepto_t() As String

    concepto_t = "feriados.concepto"

End Property

