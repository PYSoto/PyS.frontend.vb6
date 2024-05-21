VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVDAOAlterArtProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vCodigo As String
Private vDescripcion As String
Private vRazon As String

Private clon As clsVDAOAlterArtProv

Public ddl As New clsVDDLAlterArtProv

Private Sub class_Initialize()

    initializeOther

End Sub

Private Sub initializeOther()

    vClave = 0
    vCodigo = ""
    vDescripcion = ""
    vRazon = ""

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

Public Property Get codigo() As Variant

    codigo = vCodigo

End Property

Public Property Let codigo(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigo = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get razon() As Variant

    razon = vRazon

End Property

Public Property Let razon(ByVal vNewValue As Variant)

On Error Resume Next

    vRazon = vNewValue

End Property

Public Function clone() As clsVDAOAlterArtProv
Dim objObjeto As New clsVDAOAlterArtProv

    With objObjeto
        .clave = vClave
        .codigo = vCodigo
        .descripcion = vDescripcion
        .razon = vRazon
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsVDAOAlterArtProv, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .codigo = pQuery.Fields(.ddl.codigo)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .razon = pQuery.Fields(.ddl.razon)
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
    strSQL = strSQL & " WHERE " & ddl.codigo & " = '" & pCodigo & "'"
    strSQL = strSQL & ";"

    Set collectionByCodigo = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsVDAOAlterArtProv

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsVDAOAlterArtProv
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Sub llenarGrilla(pGrilla As MSFlexGrid, pArtID As String, pDB As clsDB)
Dim objAAP As clsVDAOAlterArtProv

    pGrilla.Rows = 1
    pGrilla.Redraw = False
    For Each objAAP In collectionByCodigo(pArtID, pDB)
        With objAAP
            pGrilla.AddItem modGrid.array2itemGrid(Array(.clave, .codigo, .descripcion, .razon))
        End With
    Next
    pGrilla.Redraw = True
    
End Sub

