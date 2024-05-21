VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLDetpresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "detpresupuesto"

End Property

Public Property Get preid() As String

    preid = "dep_pre_id"

End Property

Public Property Get preid_t() As String

    preid_t = "detpresupuesto.dep_pre_id"

End Property

Public Property Get orden() As String

    orden = "dep_orden"

End Property

Public Property Get orden_t() As String

    orden_t = "detpresupuesto.dep_orden"

End Property

Public Property Get artid() As String

    artid = "dep_art_id"

End Property

Public Property Get artid_t() As String

    artid_t = "detpresupuesto.dep_art_id"

End Property

Public Property Get cantidad() As String

    cantidad = "dep_cant_art"

End Property

Public Property Get cantidad_t() As String

    cantidad_t = "detpresupuesto.dep_cant_art"

End Property

Public Property Get unitsiva() As String

    unitsiva = "dep_unitsiva"

End Property

Public Property Get unitsiva_t() As String

    unitsiva_t = "detpresupuesto.dep_unitsiva"

End Property

Public Property Get unitciva() As String

    unitciva = "dep_unitciva"

End Property

Public Property Get unitciva_t() As String

    unitciva_t = "detpresupuesto.dep_unitciva"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "detpresupuesto.clave"

End Property

