VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDDLArticulosMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "articulosmirror"

End Property

Public Property Get dsn() As String

    dsn = "arm_dsn"

End Property

Public Property Get dsn_t() As String

    dsn_t = "articulosmirror.arm_dsn"

End Property

Public Property Get uid() As String

    uid = "arm_uid"

End Property

Public Property Get uid_t() As String

    uid_t = "articulosmirror.arm_uid"

End Property

Public Property Get pwd() As String

    pwd = "arm_pwd"

End Property

Public Property Get pwd_t() As String

    pwd_t = "articulosmirror.arm_pwd"

End Property

Public Property Get ip() As String

    ip = "arm_ip"

End Property

Public Property Get ip_t() As String

    ip_t = "articulosmirror.arm_ip"

End Property

Public Property Get database() As String

    database = "arm_database"

End Property

Public Property Get database_t() As String

    database_t = "articulosmirror.arm_database"

End Property

Public Property Get id() As String

    id = "arm_id"

End Property

Public Property Get id_t() As String

    id_t = "articulosmirror.arm_id"

End Property

