VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVDAOAliasProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vArtid As String
Private vAlias As String
Private vRazon As String
Private vPreciocompra As Currency

Private clon As clsVDAOAliasProv

Public ddl As New clsVDDLAliasProv

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vClave = 0
    vArtid = ""
    vAlias = ""
    vRazon = ""
    vPreciocompra = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get artid() As Variant

    artid = vArtid

End Property

Public Property Let artid(ByVal vNewValue As Variant)

On Error Resume Next

    vArtid = vNewValue

End Property

Public Property Get alias() As Variant

    alias = vAlias

End Property

Public Property Let alias(ByVal vNewValue As Variant)

On Error Resume Next

    vAlias = vNewValue

End Property

Public Property Get razon() As Variant

    razon = vRazon

End Property

Public Property Let razon(ByVal vNewValue As Variant)

On Error Resume Next

    vRazon = vNewValue

End Property

Public Property Get preciocompra() As Variant

    preciocompra = vPreciocompra

End Property

Public Property Let preciocompra(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciocompra = vNewValue

End Property

Public Function clone() As clsVDAOAliasProv
Dim objObjeto As New clsVDAOAliasProv

    With objObjeto
        .clave = vClave
        .artid = vArtid
        .alias = vAlias
        .razon = vRazon
        .preciocompra = vPreciocompra
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVDAOAliasProv, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .artid = pQuery.Fields(.ddl.artid)
        .alias = pQuery.Fields(.ddl.alias)
        .razon = pQuery.Fields(.ddl.razon)
        .preciocompra = pQuery.Fields(.ddl.preciocompra)
    End With

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByCodigo(pCodigo As String, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.artid & " = '" & pCodigo & "'"
    strSQL = strSQL & ";"

    Set collectionByCodigo = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsVDAOAliasProv

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsVDAOAliasProv
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Sub llenarGrilla(pGrilla As MSFlexGrid, pArtID As String, pDB As clsDB)
Dim objAPr As clsVDAOAliasProv

    pGrilla.Rows = 1
    pGrilla.Redraw = False
    For Each objAPr In collectionByCodigo(pArtID, pDB)
        With objAPr
            pGrilla.AddItem modGrid.array2itemGrid(Array(.clave, .alias, .razon, Format(.preciocompra, "0.00")))
        End With
    Next
    pGrilla.Redraw = True
    
End Sub

