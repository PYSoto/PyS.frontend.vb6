VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "usuarios"

End Property

Public Property Get clave() As String

    clave = "clave"

End Property

Public Property Get clave_t() As String

    clave_t = "usuarios.clave"

End Property

Public Property Get nombre() As String

    nombre = "nombre"

End Property

Public Property Get nombre_t() As String

    nombre_t = "usuarios.nombre"

End Property

Public Property Get password() As String

    password = "usu_password"

End Property

Public Property Get password_t() As String

    password_t = "usuarios.usu_password"

End Property

Public Property Get pin() As String

    pin = "usu_pin"

End Property

Public Property Get pin_t() As String

    pin_t = "usuarios.usu_pin"

End Property

