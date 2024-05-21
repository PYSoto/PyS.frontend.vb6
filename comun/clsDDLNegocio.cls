VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "negocio"

End Property

Public Property Get negID() As String

    negID = "neg_id"

End Property

Public Property Get negID_t() As String

    negID_t = "negocio.neg_id"

End Property

Public Property Get nombre() As String

    nombre = "neg_nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "negocio.neg_nombre"

End Property

Public Property Get domicilio() As String

    domicilio = "neg_domicilio"

End Property

Public Property Get domicilio_t() As String

    domicilio_t = "negocio.neg_domicilio"

End Property

Public Property Get dsn() As String

    dsn = "neg_dsn"

End Property

Public Property Get dsn_t() As String

    dsn_t = "negocio.neg_dsn"

End Property

Public Property Get ip() As String

    ip = "neg_ip"

End Property

Public Property Get ip_t() As String

    ip_t = "negocio.neg_ip"

End Property

Public Property Get db() As String

    db = "neg_db"

End Property

Public Property Get db_t() As String

    db_t = "negocio.neg_db"

End Property

Public Property Get user() As String

    user = "neg_user"

End Property

Public Property Get user_t() As String

    user_t = "negocio.neg_user"

End Property

