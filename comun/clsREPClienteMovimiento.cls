VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsREPClienteMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function findREST(json_object As Dictionary) As clsMODClienteMovimiento
Dim clienteMovimiento As New clsMODClienteMovimiento

    If json_object.item("status") <> "" Then
        Set findREST = New clsMODClienteMovimiento
        Exit Function
    End If

    clienteMovimiento.fillObject json_object

    Set findREST = clienteMovimiento

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function findByClientemovimientoId(ByVal clientemovimientoId As Long) As clsMODClienteMovimiento
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/" & clientemovimientoId

    request.Open "GET", url
    request.send

    Set findByClientemovimientoId = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function deleteByclientemovimientoId(clientemovimientoId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    deleteByclientemovimientoId = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/" & clientemovimientoId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then deleteByclientemovimientoId = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(clienteMovimiento As clsMODClienteMovimiento) As clsMODClienteMovimiento

    If exist(clienteMovimiento.clientemovimientoId) Then
        Set save = update(clienteMovimiento, clienteMovimiento.clientemovimientoId)
    Else
        Set save = add(clienteMovimiento)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(clientemovimientoId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/" & clientemovimientoId

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(clienteMovimiento As clsMODClienteMovimiento) As clsMODClienteMovimiento
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODClienteMovimiento

    Set add = New clsMODClienteMovimiento

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(clienteMovimiento.makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set add = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(clienteMovimiento As clsMODClienteMovimiento, clientemovimientoId As Long) As clsMODClienteMovimiento
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODClienteMovimiento

    Set update = clienteMovimiento

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/" & clientemovimientoId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(clienteMovimiento.makeParams(True))

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set update = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAsociables(clienteId As Long) As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "clientemovimiento/asociable/" & clienteId

    request.Open "GET", url
    request.send

    Set collectionAsociables = collectionJSON(request.responseText)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 27/5/2021
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String) As Collection
Dim elemento As clsMODClienteMovimiento

Dim registros As Collection

Dim json_object As Dictionary

    If Left(jsontext, 1) <> "[" Or jsontext = "[]" Then
        Set collectionJSON = New Collection
        Exit Function
    End If

    Set registros = New Collection

    For Each json_object In modJsonConverter.ParseJson(jsontext)
        Set elemento = New clsMODClienteMovimiento
        elemento.fillObject json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

