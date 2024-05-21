VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsREPComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Consulta la DB con un SQL buscando 1 registro
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Function findREST(json_object As Dictionary) As clsMODComprobante
Dim comprobante As New clsMODComprobante

    If json_object.item("status") <> "" Then
        Set findREST = New clsMODComprobante
        Exit Function
    End If

    comprobante.fillObject json_object

    Set findREST = comprobante

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function findByComprobanteId(comprobanteId As Integer) As clsMODComprobante
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/" & comprobanteId

    request.Open "GET", url
    request.send

    Set findByComprobanteId = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function deleteBycomprobanteId(comprobanteId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    deleteBycomprobanteId = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/" & comprobanteId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then deleteBycomprobanteId = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(comprobante As clsMODComprobante) As clsMODComprobante

    If exist(comprobante.comprobanteId) Then
        Set save = update(comprobante, comprobante.comprobanteId)
    Else
        Set save = add(comprobante)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(comprobanteId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/" & comprobanteId

    request.Open "GET", url
    request.send

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    If response_object.item("status") = "" Then exist = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Hace 1 INSERT con los valores de los atributos
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function add(comprobante As clsMODComprobante) As clsMODComprobante
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODComprobante

    Set add = New clsMODComprobante

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(comprobante.makeParams)

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set add = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function update(comprobante As clsMODComprobante, comprobanteId As Long) As clsMODComprobante
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODComprobante

    Set update = comprobante

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/" & comprobanteId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(comprobante.makeParams(True))

    request.send json_text

    Set response_object = modJsonConverter.ParseJson(request.responseText)

    object.fillObject response_object

    If response_object.item("status") = "" Then Set update = object

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Arma una colección con todos los registros
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionAll() As Collection
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "comprobante/"

    request.Open "GET", url
    request.send

    Set collectionAll = collectionJSON(request.responseText)

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Hace la consulta SQL sobre la DB y arma la colección
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function collectionJSON(jsontext As String) As Collection
Dim elemento As clsMODComprobante

Dim registros As Collection

Dim json_object As Dictionary

    If Left(jsontext, 1) <> "[" Or jsontext = "[]" Then
        Set collectionJSON = New Collection
        Exit Function
    End If

    Set registros = New Collection

    For Each json_object In modJsonConverter.ParseJson(jsontext)
        Set elemento = New clsMODComprobante
        elemento.fillObject json_object

        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key
    Next

    Set collectionJSON = registros

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Llena una ComboBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillCombo(pCombo As ComboBox)

    fillComboCollection pCombo, collectionAll()

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Llena la ComboBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)
Dim elemento As clsMODComprobante

    pCombo.Clear

    For Each elemento In pCollection
        pCombo.AddItem elemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Llena una ListBox con todos los datos de la tabla
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Sub fillList(pList As ListBox)

    fillListCollection pList, collectionAll()

End Sub

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Llena la ListBox efectivamente
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Private Sub fillListCollection(pList As ListBox, pCollection As Collection)
Dim elemento As clsMODComprobante

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

