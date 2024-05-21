VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "empresa"

End Property

Public Property Get id() As String

    id = "emp_id"

End Property

Public Property Get id_t() As String

    id_t = "empresa.emp_id"

End Property

Public Property Get negID() As String

    negID = "emp_neg_id"

End Property

Public Property Get negID_t() As String

    negID_t = "empresa.emp_neg_id"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "empresa.nombre"

End Property

Public Property Get rsocial() As String

    rsocial = "emp_rsocial"

End Property

Public Property Get rsocial_t() As String

    rsocial_t = "empresa.emp_rsocial"

End Property

Public Property Get domicilio() As String

    domicilio = "domicilio"

End Property

Public Property Get domicilio_t() As String

    domicilio_t = "empresa.domicilio"

End Property

Public Property Get telf() As String

    telf = "telf"

End Property

Public Property Get telf_t() As String

    telf_t = "empresa.telf"

End Property

Public Property Get cuit() As String

    cuit = "cuit"

End Property

Public Property Get cuit_t() As String

    cuit_t = "empresa.cuit"

End Property

Public Property Get puntovta() As String

    puntovta = "puntovta"

End Property

Public Property Get puntovta_t() As String

    puntovta_t = "empresa.puntovta"

End Property

Public Property Get ingbrutos() As String

    ingbrutos = "ingbrutos"

End Property

Public Property Get ingbrutos_t() As String

    ingbrutos_t = "empresa.ingbrutos"

End Property

Public Property Get nroestablecimiento() As String

    nroestablecimiento = "nroestablecimiento"

End Property

Public Property Get nroestablecimiento_t() As String

    nroestablecimiento_t = "empresa.nroestablecimiento"

End Property

Public Property Get sedetimbrado() As String

    sedetimbrado = "sedetimbrado"

End Property

Public Property Get sedetimbrado_t() As String

    sedetimbrado_t = "empresa.sedetimbrado"

End Property

Public Property Get inicioactividades() As String

    inicioactividades = "inicioactividades"

End Property

Public Property Get inicioactividades_t() As String

    inicioactividades_t = "empresa.inicioactividades"

End Property

Public Property Get condicioniva() As String

    condicioniva = "condicioniva"

End Property

Public Property Get condicioniva_t() As String

    condicioniva_t = "empresa.condicioniva"

End Property

Public Property Get ubicacion() As String

    ubicacion = "emp_ubicacion"

End Property

Public Property Get ubicacion_t() As String

    ubicacion_t = "empresa.emp_ubicacion"

End Property


