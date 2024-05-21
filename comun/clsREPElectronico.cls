VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsREPElectronico"
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
Private Function findREST(json_object As Dictionary) As clsMODElectronico
Dim electronico As New clsMODElectronico

    If json_object.item("status") <> "" Then
        Set findREST = New clsMODElectronico
        Exit Function
    End If

    electronico.fillObject json_object

    Set findREST = electronico

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Realiza un find sobre los atributos de la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function findByComprobante(comprobanteId As Integer, puntoventa As Integer, numerocomprobante As Long) As clsMODElectronico
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "electronico/comprobante" & modUrls.join(Array(comprobanteId, puntoventa, numerocomprobante))

    request.Open "GET", url
    request.send

    Set findByComprobante = findREST(modJsonConverter.ParseJson(request.responseText))

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Elimina 1 registro de la DB según la Primary Key
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function deleteByelectronicoId(electronicoId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

    deleteByelectronicoId = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "electronico/" & electronicoId

    request.Open "DELETE", url
    request.send

    If request.responseText = "" Then deleteByelectronicoId = True

    Set request = Nothing

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Graba un registro verificando si debe agregarlo o actualizarlo
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function save(electronico As clsMODElectronico) As clsMODElectronico

    If exist(electronico.electronicoId) Then
        Set save = update(electronico, electronico.electronicoId)
    Else
        Set save = add(electronico)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * *
' Author: Ing. Daniel Quinteros
' Created: 28/5/2021
' Description: Verifica si el registro asociado a la Primary Key existe
' Log:
' * * * * * * * * * * * * * * * * * * * * *
Public Function exist(electronicoId As Long) As Boolean
Dim request As MSXML2.ServerXMLHTTP

Dim url As String

Dim response_object As Dictionary

    exist = False

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "electronico/" & electronicoId

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
Public Function add(electronico As clsMODElectronico) As clsMODElectronico
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODElectronico

    Set add = New clsMODElectronico

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "electronico/"

    request.Open "POST", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(electronico.makeParams)

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
Public Function update(electronico As clsMODElectronico, electronicoId As Long) As clsMODElectronico
Dim request As MSXML2.ServerXMLHTTP

Dim url As String
Dim json_text As String

Dim response_object As Dictionary

Dim object As New clsMODElectronico

    Set update = electronico

    Set request = New MSXML2.ServerXMLHTTP

    url = modUrls.url_pys() & "electronico/" & electronicoId

    request.Open "PUT", url
    request.setRequestHeader "Content-Type", "application/json"

    json_text = modJsonConverter.ConvertToJson(electronico.makeParams(True))

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

    url = modUrls.url_pys() & "electronico/"

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
Dim elemento As clsMODElectronico

Dim registros As Collection

Dim json_object As Dictionary

    If Left(jsontext, 1) <> "[" Or jsontext = "[]" Then
        Set collectionJSON = New Collection
        Exit Function
    End If

    Set registros = New Collection

    For Each json_object In modJsonConverter.ParseJson(jsontext)
        Set elemento = New clsMODElectronico
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
Dim elemento As clsMODElectronico

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
Dim elemento As clsMODElectronico

    pList.Clear

    For Each elemento In pCollection
        pList.AddItem elemento.listText
        pList.ItemData(pList.NewIndex) = elemento.listKey
    Next

End Sub

