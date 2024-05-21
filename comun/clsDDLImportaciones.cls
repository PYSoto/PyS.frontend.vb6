VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLImportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "importaciones"

End Property

Public Property Get impID() As String

    impID = "imp_id"

End Property

Public Property Get impID_t() As String

    impID_t = "importaciones.imp_id"

End Property

Public Property Get fecha() As String

    fecha = "imp_fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "importaciones.imp_fecha"

End Property

