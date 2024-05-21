VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOProveedorPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vProveedorMovimientoIDDeuda As Long
Private vProveedorMovimientoIDAplicado As Long
Private vImporteAplicado As Currency
Private vAutoID As Long
Private vCreated As Date

Private clon As clsDAOProveedorPago

Public ddl As New clsDDLProveedorPago

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de la Clase
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub class_Initialize()

    vProveedorMovimientoIDDeuda = 0
    vProveedorMovimientoIDAplicado = 0

    initializeOther

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Inicializacion de los campos que no corresponden a la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub initializeOther()

    vImporteAplicado = 0
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
' Description: Getter proveedorMovimientoIDDeuda
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDDeuda() As Variant

    proveedorMovimientoIDDeuda = vProveedorMovimientoIDDeuda

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedorMovimientoIDDeuda
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedorMovimientoIDDeuda(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedorMovimientoIDDeuda = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter proveedorMovimientoIDAplicado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get proveedorMovimientoIDAplicado() As Variant

    proveedorMovimientoIDAplicado = vProveedorMovimientoIDAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter proveedorMovimientoIDAplicado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let proveedorMovimientoIDAplicado(ByVal vNewValue As Variant)

On Error Resume Next

    vProveedorMovimientoIDAplicado = vNewValue

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter importeAplicado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get importeAplicado() As Variant

    importeAplicado = vImporteAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Setter importeAplicado
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Let importeAplicado(ByVal vNewValue As Variant)

On Error Resume Next

    vImporteAplicado = vNewValue

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

Public Property Get key() As String

    key = "k." & autoID
    
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

    comboText = vImporteAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter comboKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get comboKey() As Variant

    comboKey = vProveedorMovimientoIDAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listText
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listText() As Variant

    listText = vImporteAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Getter listKey
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Property Get listKey() As Variant

    listKey = vProveedorMovimientoIDAplicado

End Property

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Clona el objeto this y devuelve uno nuevo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function clone() As clsDAOProveedorPago
Dim objObjeto As New clsDAOProveedorPago

    With objObjeto
        .proveedorMovimientoIDDeuda = vProveedorMovimientoIDDeuda
        .proveedorMovimientoIDAplicado = vProveedorMovimientoIDAplicado
        .importeAplicado = vImporteAplicado
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
Private Sub fillObject(pObjeto As clsDAOProveedorPago, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .proveedorMovimientoIDDeuda = pQuery.Fields(.ddl.proveedorMovimientoIDDeuda)
        .proveedorMovimientoIDAplicado = pQuery.Fields(.ddl.proveedorMovimientoIDAplicado)
        .importeAplicado = pQuery.Fields(.ddl.importeAplicado)
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
    sql = sql & " WHERE " & ddl.proveedorMovimientoIDDeuda & " = " & vProveedorMovimientoIDDeuda
    sql = sql & " AND " & ddl.proveedorMovimientoIDAplicado & " = " & vProveedorMovimientoIDAplicado

    makeWherePrimaryKey = sql

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB con un SQL buscando 1 registro
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
' Description: Consulta la pDB con un SQL buscando el último registro según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub findLast(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedorMovimientoIDDeuda & " = " & vProveedorMovimientoIDDeuda
    sql = sql & " ORDER BY " & ddl.proveedorMovimientoIDAplicado & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Consulta la pDB con un SQL buscando 1 registro de acuerdo a los atributos cargados
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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Elimina 1 registro de la pDB según la Primary Key
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

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(pDB As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.importeAplicado & " = '" & Replace(vImporteAplicado, ",", ".") & "'"
    sql = sql & ", " & ddl.proveedorMovimientoIDDeuda & " = " & vProveedorMovimientoIDDeuda
    sql = sql & ", " & ddl.proveedorMovimientoIDAplicado & " = " & vProveedorMovimientoIDAplicado
    sql = sql & ";"

    add = pDB.execute(sql)

    findByPrimaryKey pDB

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la pDB
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

    If vImporteAplicado <> clon.importeAplicado Or IsNull(vImporteAplicado) <> IsNull(clon.importeAplicado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importeAplicado & " = '" & Replace(vImporteAplicado, ",", ".") & "'"
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

    vProveedorMovimientoIDAplicado = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedorMovimientoIDDeuda & " = " & vProveedorMovimientoIDDeuda
    sql = sql & " ORDER BY " & ddl.proveedorMovimientoIDAplicado & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vProveedorMovimientoIDAplicado = 1 + rstQuery.Fields(ddl.proveedorMovimientoIDAplicado)

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

Public Function collectionByPago(proveedormovimientoID As Long, pDB As clsDB) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.proveedorMovimientoIDAplicado & " = " & proveedormovimientoID
    sql = sql & ";"

    Set collectionByPago = collectionSQL(sql, pDB)

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 01/05/2018
' Description: Hace la consulta SQL sobre la pDB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionSQL(sql As String, pDB As clsDB, Optional registrosold As Collection = Nothing) As Collection
Dim query As ADODB.Recordset

Dim elemento As clsDAOProveedorPago

Dim registros As Collection

    If registrosold Is Nothing Then
        Set registros = New Collection
    Else
        Set registros = registrosold
    End If

    Set query = pDB.query(sql)

    Do While Not query.EOF
        Set elemento = New clsDAOProveedorPago
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
Dim elemento As clsDAOProveedorPago

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
Dim elemento As clsDAOProveedorPago

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

