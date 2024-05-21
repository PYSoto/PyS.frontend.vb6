VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLConsfecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "consfecha"

End Property

Public Property Get negid() As String

    negid = "cfe_neg_id"

End Property

Public Property Get negid_t() As String

    negid_t = "consfecha.cfe_neg_id"

End Property

Public Property Get fecha() As String

    fecha = "cfe_fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "consfecha.cfe_fecha"

End Property

Public Property Get cuenta() As String

    cuenta = "cfe_cuenta"

End Property

Public Property Get cuenta_t() As String

    cuenta_t = "consfecha.cfe_cuenta"

End Property

Public Property Get deudor() As String

    deudor = "cfe_deudor"

End Property

Public Property Get deudor_t() As String

    deudor_t = "consfecha.cfe_deudor"

End Property

Public Property Get acreedor() As String

    acreedor = "cfe_acreedor"

End Property

Public Property Get acreedor_t() As String

    acreedor_t = "consfecha.cfe_acreedor"

End Property

