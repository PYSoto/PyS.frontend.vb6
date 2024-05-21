VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDLPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get table() As String

    table = "presupuesto"

End Property

Public Property Get preID() As String

    preID = "pre_id"

End Property

Public Property Get preid_t() As String

    preid_t = "presupuesto.pre_id"

End Property

Public Property Get cliID() As String

    cliID = "pre_cli_id"

End Property

Public Property Get cliID_t() As String

    cliID_t = "presupuesto.pre_cli_id"

End Property

Public Property Get fecha() As String

    fecha = "pre_fecha"

End Property

Public Property Get fecha_t() As String

    fecha_t = "presupuesto.pre_fecha"

End Property

Public Property Get fechavto() As String

    fechavto = "pre_fechavto"

End Property

Public Property Get fechavto_t() As String

    fechavto_t = "presupuesto.pre_fechavto"

End Property

Public Property Get observac() As String

    observac = "pre_observac"

End Property

Public Property Get observac_t() As String

    observac_t = "presupuesto.pre_observac"

End Property

Public Property Get clavemovclie() As String

    clavemovclie = "pre_clavemovclie"

End Property

Public Property Get clavemovclie_t() As String

    clavemovclie_t = "presupuesto.pre_clavemovclie"

End Property

Public Property Get ctacte() As String

    ctacte = "pre_ctacte"

End Property

Public Property Get ctacte_t() As String

    ctacte_t = "presupuesto.pre_ctacte"

End Property


