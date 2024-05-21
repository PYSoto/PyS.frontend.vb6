VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLTMBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "tmbalance"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "tmbalance.clave"

End Property

Public Property Get hwnd() As String

    hwnd = "bal_hwnd"

End Property

Public Property Get hwnd_t() As String

    hwnd_t = "tmbalance.bal_hwnd"

End Property

Public Property Get cuenta() As String

    cuenta = "bal_cuenta"

End Property

Public Property Get cuenta_t() As String

    cuenta_t = "tmbalance.bal_cuenta"

End Property

Public Property Get deudor() As String

    deudor = "bal_deudor"

End Property

Public Property Get deudor_t() As String

    deudor_t = "tmbalance.bal_deudor"

End Property

Public Property Get acreedor() As String

    acreedor = "bal_acreedor"

End Property

Public Property Get acreedor_t() As String

    acreedor_t = "tmbalance.bal_acreedor"

End Property

Public Property Get saldodeudor() As String

    saldodeudor = "bal_saldodeudor"

End Property

Public Property Get saldodeudor_t() As String

    saldodeudor_t = "tmbalance.bal_saldodeudor"

End Property

Public Property Get saldoacreedor() As String

    saldoacreedor = "bal_saldoacreedor"

End Property

Public Property Get saldoacreedor_t() As String

    saldoacreedor_t = "tmbalance.bal_saldoacreedor"

End Property

Public Property Get grado() As String

    grado = "bal_grado"

End Property

Public Property Get grado_t() As String

    grado_t = "tmbalance.bal_grado"

End Property

