VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vProveedorID As Long
Private vRazonSocial As String
Private vNombreFantasia As String
Private vCuit As String
Private vDomicilio As String
Private vLocalidad As String
Private vProvincia As String
Private vTelefono As String
Private vFax As String
Private vEmail As String
Private vPosicionIva As Integer
Private vCelular As String
Private vIngresosBrutos As String
Private vContacto As String
Private vObservaciones As String
Private vAutoID As Long
Private vCreated As Date

Private clon As clsDAOProveedor

Public ddl As New clsDDLProveedor

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vProveedorID = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vRazonSocial = ""
    vNombreFantasia = ""
    vCuit = ""
    vDomicilio = ""
    vLocalidad = ""
    vProvincia = ""
    vTelefono = ""
    vFax = ""
    vEmail = ""
    vPosicionIva = 0
    vCelular = ""
    vIngresosBrutos = ""
    vContacto = ""
    vObservaciones = ""
    vAutoID = 0
    vCreated = Date

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Limpia todos los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub clean()

    Call class_Initialize

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter proveedorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorID() As Variant

    proveedorID = vProveedorID

End Property

Public Property Get key() As String

    key = "k." & proveedorID
    
End Property

Public Property Get idSearch() As Integer

    idSearch = proveedorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedorID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedorID(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedorID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter razonSocial
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get razonSocial() As Variant

    razonSocial = vRazonSocial

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter razonSocial
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let razonSocial(ByVal vNewValue As Variant)

On Error Resume Next

    vRazonSocial = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter nombreFantasia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get nombreFantasia() As Variant

    nombreFantasia = vNombreFantasia

End Property

Public Property Get textSearch() As String

    textSearch = vRazonSocial & " - " & vNombreFantasia & " (" & vProveedorID & ")"
    
End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter nombreFantasia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let nombreFantasia(ByVal vNewValue As Variant)

On Error Resume Next

    vNombreFantasia = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter cuit
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get cuit() As Variant

    cuit = vCuit

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter cuit
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter domicilio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get domicilio() As Variant

    domicilio = vDomicilio

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter domicilio
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let domicilio(ByVal vNewValue As Variant)

On Error Resume Next

    vDomicilio = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter localidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get localidad() As Variant

    localidad = vLocalidad

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter localidad
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let localidad(ByVal vNewValue As Variant)

On Error Resume Next

    vLocalidad = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter provincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get provincia() As Variant

    provincia = vProvincia

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter provincia
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let provincia(ByVal vNewValue As Variant)

On Error Resume Next

    vProvincia = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter telefono
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get telefono() As Variant

    telefono = vTelefono

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter telefono
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let telefono(ByVal vNewValue As Variant)

On Error Resume Next

    vTelefono = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter fax
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get fax() As Variant

    fax = vFax

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter fax
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let fax(ByVal vNewValue As Variant)

On Error Resume Next

    vFax = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter email
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get email() As Variant

    email = vEmail

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter email
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let email(ByVal vNewValue As Variant)

On Error Resume Next

    vEmail = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter posicionIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get posicionIva() As Variant

    posicionIva = vPosicionIva

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter posicionIva
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let posicionIva(ByVal vNewValue As Variant)

On Error Resume Next

    vPosicionIva = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter celular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get celular() As Variant

    celular = vCelular

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter celular
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let celular(ByVal vNewValue As Variant)

On Error Resume Next

    vCelular = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter ingresosBrutos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get ingresosBrutos() As Variant

    ingresosBrutos = vIngresosBrutos

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter ingresosBrutos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let ingresosBrutos(ByVal vNewValue As Variant)

On Error Resume Next

    vIngresosBrutos = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter contacto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get contacto() As Variant

    contacto = vContacto

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter contacto
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let contacto(ByVal vNewValue As Variant)

On Error Resume Next

    vContacto = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get observaciones() As Variant

    observaciones = vObservaciones

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter observaciones
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let observaciones(ByVal vNewValue As Variant)

On Error Resume Next

    vObservaciones = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get autoID() As Variant

    autoID = vAutoID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter autoID
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let autoID(ByVal vNewValue As Variant)

On Error Resume Next

    vAutoID = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get created() As Variant

    created = vCreated

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter created
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let created(ByVal vNewValue As Variant)

On Error Resume Next

    vCreated = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboText() As Variant

    comboText = vRazonSocial & " - " & vNombreFantasia & " (" & vProveedorID & ")"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vProveedorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vRazonSocial & " - " & vNombreFantasia & " (" & vProveedorID & ")"

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vProveedorID

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOProveedor
Dim objObjeto As New clsDAOProveedor

    With objObjeto
        .proveedorID = vProveedorID
        .razonSocial = vRazonSocial
        .nombreFantasia = vNombreFantasia
        .cuit = vCuit
        .domicilio = vDomicilio
        .localidad = vLocalidad
        .provincia = vProvincia
        .telefono = vTelefono
        .fax = vFax
        .email = vEmail
        .posicionIva = vPosicionIva
        .celular = vCelular
        .ingresosBrutos = vIngresosBrutos
        .contacto = vContacto
        .observaciones = vObservaciones
        .autoID = vAutoID
        .created = vCreated
    End With

    Set clone = objObjeto

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Carga los campos del registro en los atributos de la clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillObject(pObjeto As clsDAOProveedor, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .proveedorID = pQuery.Fields(.ddl.proveedorID)
        .razonSocial = pQuery.Fields(.ddl.razonSocial)
        .nombreFantasia = pQuery.Fields(.ddl.nombreFantasia)
        .cuit = pQuery.Fields(.ddl.cuit)
        .domicilio = pQuery.Fields(.ddl.domicilio)
        .localidad = pQuery.Fields(.ddl.localidad)
        .provincia = pQuery.Fields(.ddl.provincia)
        .telefono = pQuery.Fields(.ddl.telefono)
        .fax = pQuery.Fields(.ddl.fax)
        .email = pQuery.Fields(.ddl.email)
        .posicionIva = pQuery.Fields(.ddl.posicionIva)
        .celular = pQuery.Fields(.ddl.celular)
        .ingresosBrutos = pQuery.Fields(.ddl.ingresosBrutos)
        .contacto = pQuery.Fields(.ddl.contacto)
        .observaciones = pQuery.Fields(.ddl.observaciones)
        .autoID = pQuery.Fields(.ddl.autoID)
        .created = pQuery.Fields(.ddl.created)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Construye el SQL de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.proveedorID & " = " & vProveedorID

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub findSQL(pSQL As String, pDB As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = pDB.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByPrimaryKey(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.proveedorID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la DB con un SQL buscando 1 registro de acuerdo a los atributos cargados
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findByAutoID(pDB As clsDB)
Dim sql As String

Dim lngAutoID As Long

    lngAutoID = vAutoID

    clean

    vAutoID = lngAutoID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.autoID & " = " & vAutoID
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Sub findSearch(proveedorID As Integer, db As clsDB)

    vProveedorID = proveedorID
    findByPrimaryKey db
    
End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function delete(pDB As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete = pDB.execute(sql)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(pDB As clsDB) As Boolean

    If exist(pDB) Then
        save = update(pDB)
    Else
        save = add(pDB)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(pDB As clsDB) As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function existByCUIT(db As clsDB, Optional nombre As String = "") As Boolean
Dim sql As String

Dim rstQuery As ADODB.Recordset

    If vCuit = "00-00000000-0" Then
        existByCUIT = False
        Exit Function
    End If

    nombre = ""

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.cuit & " = '" & vCuit & "'"
    sql = sql & " AND " & ddl.proveedorID & " <> " & vProveedorID
    sql = sql & ";"
    
    Set rstQuery = db.query(sql)
    existByCUIT = Not rstQuery.EOF
    If Not rstQuery.EOF Then nombre = rstQuery.Fields(ddl.razonSocial)
    rstQuery.Close
    
End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(pDB As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.razonSocial & " = '" & replaceSpecialSymbols(vRazonSocial) & "'"
    sql = sql & ", " & ddl.nombreFantasia & " = '" & replaceSpecialSymbols(vNombreFantasia) & "'"
    sql = sql & ", " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
    sql = sql & ", " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
    sql = sql & ", " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
    sql = sql & ", " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
    sql = sql & ", " & ddl.telefono & " = '" & replaceSpecialSymbols(vTelefono) & "'"
    sql = sql & ", " & ddl.fax & " = '" & replaceSpecialSymbols(vFax) & "'"
    sql = sql & ", " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
    sql = sql & ", " & ddl.posicionIva & " = " & vPosicionIva
    sql = sql & ", " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
    sql = sql & ", " & ddl.ingresosBrutos & " = '" & replaceSpecialSymbols(vIngresosBrutos) & "'"
    sql = sql & ", " & ddl.contacto & " = '" & replaceSpecialSymbols(vContacto) & "'"
    sql = sql & ", " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
    sql = sql & ", " & ddl.proveedorID & " = " & vProveedorID
    sql = sql & ";"

    add = pDB.execute(sql)

    findByPrimaryKey pDB

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(pDB As clsDB) As Boolean
Dim sql As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vRazonSocial <> clon.razonSocial Or IsNull(vRazonSocial) <> IsNull(clon.razonSocial) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.razonSocial & " = '" & replaceSpecialSymbols(vRazonSocial) & "'"
        blnChange = True
    End If
    If vNombreFantasia <> clon.nombreFantasia Or IsNull(vNombreFantasia) <> IsNull(clon.nombreFantasia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombreFantasia & " = '" & replaceSpecialSymbols(vNombreFantasia) & "'"
        blnChange = True
    End If
    If vCuit <> clon.cuit Or IsNull(vCuit) <> IsNull(clon.cuit) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
        blnChange = True
    End If
    If vDomicilio <> clon.domicilio Or IsNull(vDomicilio) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
        blnChange = True
    End If
    If vLocalidad <> clon.localidad Or IsNull(vLocalidad) <> IsNull(clon.localidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
        blnChange = True
    End If
    If vProvincia <> clon.provincia Or IsNull(vProvincia) <> IsNull(clon.provincia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
        blnChange = True
    End If
    If vTelefono <> clon.telefono Or IsNull(vTelefono) <> IsNull(clon.telefono) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.telefono & " = '" & replaceSpecialSymbols(vTelefono) & "'"
        blnChange = True
    End If
    If vFax <> clon.fax Or IsNull(vFax) <> IsNull(clon.fax) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fax & " = '" & replaceSpecialSymbols(vFax) & "'"
        blnChange = True
    End If
    If vEmail <> clon.email Or IsNull(vEmail) <> IsNull(clon.email) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
        blnChange = True
    End If
    If vPosicionIva <> clon.posicionIva Or IsNull(vPosicionIva) <> IsNull(clon.posicionIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.posicionIva & " = " & vPosicionIva
        blnChange = True
    End If
    If vCelular <> clon.celular Or IsNull(vCelular) <> IsNull(clon.celular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
        blnChange = True
    End If
    If vIngresosBrutos <> clon.ingresosBrutos Or IsNull(vIngresosBrutos) <> IsNull(clon.ingresosBrutos) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ingresosBrutos & " = '" & replaceSpecialSymbols(vIngresosBrutos) & "'"
        blnChange = True
    End If
    If vContacto <> clon.contacto Or IsNull(vContacto) <> IsNull(clon.contacto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.contacto & " = '" & replaceSpecialSymbols(vContacto) & "'"
        blnChange = True
    End If
    If vObservaciones <> clon.observaciones Or IsNull(vObservaciones) <> IsNull(clon.observaciones) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update = pDB.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Busca un nuevo ID según la Primary Key y lo carga al objeto si se le indica como parámetro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub newID(pFind As Boolean, pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vProveedorID = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " ORDER BY " & ddl.proveedorID & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vProveedorID = 1 + rstQuery.Fields(ddl.proveedorID)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll(pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Arma una colección con 1 registro para verificar si algo existe en la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAny(pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set collectionAny = collectionSQL(sql, pDB)

End Function

Public Function collectionSearch(cadena As String, pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE CONCAT(" & ddl.proveedorID & ", ' ', " & ddl.razonSocial & ", ' ', " & ddl.nombreFantasia & ") LIKE '%" & cadena & "%'"
    sql = sql & " AND " & ddl.razonSocial & " != ''"
    sql = sql & " ORDER BY " & ddl.razonSocial
    sql = sql & " LIMIT 50"
    sql = sql & ";"
    
    Set collectionSearch = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOProveedor

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOProveedor
        fillObject elemento, query

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key

        query.MoveNext
    Loop

    query.Close

    Set collectionSQL = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox, pDB As clsDB)

    fillComboCollection pCombo, collectionAll(pDB)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsDAOProveedor

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox, pDB As clsDB)

    fillListCollection pList, collectionAll(pDB)

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsDAOProveedor

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

Public Sub exportar(pArchivo As String, pDB As clsDB)
Dim strSQL As String
Dim strLinea As String

Dim rstQuery As ADODB.Recordset
    
    Open pArchivo For Append As #1

    strSQL = "SELECT CONCAT('INSERT INTO " & ddl.table
    strSQL = strSQL & "( " & ddl.proveedorID
    strSQL = strSQL & ", " & ddl.cuit
    strSQL = strSQL & ", " & ddl.razonSocial
    strSQL = strSQL & ", " & ddl.domicilio
    strSQL = strSQL & ", " & ddl.localidad
    strSQL = strSQL & ", " & ddl.provincia
    strSQL = strSQL & ", " & ddl.telefono
    strSQL = strSQL & ", " & ddl.fax
    strSQL = strSQL & ", " & ddl.email
    strSQL = strSQL & ", " & ddl.posicionIva
    strSQL = strSQL & ", " & ddl.celular
    strSQL = strSQL & ", " & ddl.ingresosBrutos
    strSQL = strSQL & ", " & ddl.observaciones
    strSQL = strSQL & ") VALUES "
    strSQL = strSQL & "( ', " & ddl.proveedorID & ", '"
    strSQL = strSQL & ", ''', " & ddl.cuit & ", '''"
    strSQL = strSQL & ", ''', " & ddl.razonSocial & ", '''"
    strSQL = strSQL & ", ''', " & ddl.domicilio & ", '''"
    strSQL = strSQL & ", ''', " & ddl.localidad & ", '''"
    strSQL = strSQL & ", ''', " & ddl.provincia & ", '''"
    strSQL = strSQL & ", ''', " & ddl.telefono & ", '''"
    strSQL = strSQL & ", ''', " & ddl.fax & ", '''"
    strSQL = strSQL & ", ''', " & ddl.email & ", '''"
    strSQL = strSQL & ", ', " & ddl.posicionIva & ", '"
    strSQL = strSQL & ", ''', " & ddl.celular & ", '''"
    strSQL = strSQL & ", ''', " & ddl.ingresosBrutos & ", '''"
    strSQL = strSQL & ", ''', " & ddl.observaciones & ", '''"
    strSQL = strSQL & ");') AS cadena FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.proveedorID
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    Do While Not rstQuery.EOF
        Print #1, rstQuery!cadena
        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Close #1

End Sub

Public Sub truncate(pDB As clsDB)
Dim strSQL As String

    strSQL = "TRUNCATE " & ddl.table
    strSQL = strSQL & ";"

    pDB.execute strSQL

End Sub
