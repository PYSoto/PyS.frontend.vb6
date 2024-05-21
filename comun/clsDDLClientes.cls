VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "clientes"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "clientes.clave"

End Property

Public Property Get negID() As String

    negID = "cli_neg_id"

End Property

Public Property Get negID_t() As String

    negID_t = "clientes.cli_neg_id"

End Property

Public Property Get codigo() As String

    codigo = "codigo"

End Property

Public Property Get codigo_t() As String

    codigo_t = "clientes.codigo"

End Property

Public Property Get cuit() As String

    cuit = "cuit"

End Property

Public Property Get cuit_t() As String

    cuit_t = "clientes.cuit"

End Property

Public Property Get razon() As String

    razon = "razon"

End Property

Public Property Get razon_t() As String

    razon_t = "clientes.razon"

End Property

Public Property Get domicilio() As String

    domicilio = "domicilio"

End Property

Public Property Get domicilio_t() As String

    domicilio_t = "clientes.domicilio"

End Property

Public Property Get localidad() As String

    localidad = "cli_localidad"

End Property

Public Property Get localidad_t() As String

    localidad_t = "clientes.cli_localidad"

End Property

Public Property Get provincia() As String

    provincia = "cli_provincia"

End Property

Public Property Get provincia_t() As String

    provincia_t = "clientes.cli_provincia"

End Property

Public Property Get tel() As String

    tel = "tel"

End Property

Public Property Get tel_t() As String

    tel_t = "clientes.tel"

End Property

Public Property Get fax() As String

    fax = "fax"

End Property

Public Property Get fax_t() As String

    fax_t = "clientes.fax"

End Property

Public Property Get email() As String

    email = "email"

End Property

Public Property Get email_t() As String

    email_t = "clientes.email"

End Property

Public Property Get celular() As String

    celular = "celular"

End Property

Public Property Get celular_t() As String

    celular_t = "clientes.celular"

End Property

Public Property Get posicion() As String

    posicion = "posicion"

End Property

Public Property Get posicion_t() As String

    posicion_t = "clientes.posicion"

End Property

Public Property Get tipodoc() As String

    tipodoc = "tipodoc"

End Property

Public Property Get tipodoc_t() As String

    tipodoc_t = "clientes.tipodoc"

End Property

Public Property Get nrodoc() As String

    nrodoc = "nrodoc"

End Property

Public Property Get nrodoc_t() As String

    nrodoc_t = "clientes.nrodoc"

End Property

Public Property Get limitecredito() As String

    limitecredito = "limitecredito"

End Property

Public Property Get limitecredito_t() As String

    limitecredito_t = "clientes.limitecredito"

End Property

Public Property Get nacionalidad() As String

    nacionalidad = "nacionalidad"

End Property

Public Property Get nacionalidad_t() As String

    nacionalidad_t = "clientes.nacionalidad"

End Property

Public Property Get descuento() As String

    descuento = "cli_descuento"

End Property

Public Property Get descuento_t() As String

    descuento_t = "clientes.cli_descuento"

End Property

Public Property Get facturable() As String

    facturable = "cli_facturable"

End Property

Public Property Get facturable_t() As String

    facturable_t = "clientes.cli_facturable"

End Property

