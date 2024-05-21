VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "valores"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "valores.clave"

End Property

Public Property Get negid() As String

    negid = "val_neg_id"

End Property

Public Property Get negid_t() As String

    negid_t = "valores.val_neg_id"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = "valores.codigo"

End Property

Public Property Get cgocli() As String

    cgocli = "cgocli"

End Property

Public Property Get cgocli_t() As String

    cgocli_t = "valores.cgocli"

End Property

Public Property Get fechaemi() As String

    fechaemi = "fechaemi"

End Property

Public Property Get fechaemi_t() As String

    fechaemi_t = "valores.fechaemi"

End Property

Public Property Get fechavto() As String

    fechavto = "fechavto"

End Property

Public Property Get fechavto_t() As String

    fechavto_t = "valores.fechavto"

End Property

Public Property Get tcoid() As String

    tcoid = "val_tco_id"

End Property

Public Property Get tcoid_t() As String

    tcoid_t = "valores.val_tco_id"

End Property

Public Property Get nrocomprob() As String

    nrocomprob = "nrocomprob"

End Property

Public Property Get nrocomprob_t() As String

    nrocomprob_t = "valores.nrocomprob"

End Property

Public Property Get importe() As String

    importe = "importe"

End Property

Public Property Get importe_t() As String

    importe_t = "valores.importe"

End Property

Public Property Get fechaReg() As String

    fechaReg = "fechaReg"

End Property

Public Property Get fechaReg_t() As String

    fechaReg_t = "valores.fechaReg"

End Property

Public Property Get clavemovv() As String

    clavemovv = "clavemovv"

End Property

Public Property Get clavemovv_t() As String

    clavemovv_t = "valores.clavemovv"

End Property

Public Property Get titular() As String

    titular = "titular"

End Property

Public Property Get titular_t() As String

    titular_t = "valores.titular"

End Property

Public Property Get banco() As String

    banco = "banco"

End Property

Public Property Get banco_t() As String

    banco_t = "valores.banco"

End Property

