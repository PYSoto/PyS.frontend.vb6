VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "parametros"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "parametros.clave"

End Property

Public Property Get negid() As String

    negid = "par_neg_id"

End Property

Public Property Get negid_t() As String

    negid_t = "parametros.par_neg_id"

End Property

Public Property Get iva1() As String

    iva1 = "iva1"

End Property

Public Property Get iva1_t() As String

    iva1_t = "parametros.iva1"

End Property

Public Property Get iva2() As String

    iva2 = "iva2"

End Property

Public Property Get iva2_t() As String

    iva2_t = "parametros.iva2"

End Property

Public Property Get feproduccion() As String

    feproduccion = "par_feproduccion"

End Property

Public Property Get feproduccion_t() As String

    feproduccion_t = "parametros.par_feproduccion"

End Property

Public Property Get ta() As String

    ta = "par_ta"

End Property

Public Property Get ta_t() As String

    ta_t = "parametros.par_ta"

End Property


