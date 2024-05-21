VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOUsuarioActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vIP As String
Private vHwnd As Long
Private vLogin As String

Private clon As clsDAOUsuarioActivo

Public ddl As New clsDDLUsuarioActivo

Private Sub class_Initialize()

    vIP = ""
    vHwnd = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vLogin = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get ip() As Variant

    ip = vIP

End Property

Public Property Let ip(ByVal vNewValue As Variant)

On Error Resume Next

    vIP = vNewValue

End Property

Public Property Get hwnd() As Variant

    hwnd = vHwnd

End Property

Public Property Let hwnd(ByVal vNewValue As Variant)

On Error Resume Next

    vHwnd = vNewValue

End Property

Public Property Get login() As Variant

    login = vLogin

End Property

Public Property Let login(ByVal vNewValue As Variant)

On Error Resume Next

    vLogin = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vLogin

End Property

Public Property Get comboKey() As Variant

    comboKey = vHwnd

End Property

Public Property Get listText() As Variant

    listText = vLogin

End Property

Public Property Get listKey() As Variant

    listKey = vHwnd

End Property

Public Function clone() As clsDAOUsuarioActivo
Dim objObjeto As New clsDAOUsuarioActivo

    With objObjeto
        .ip = vIP
        .hwnd = vHwnd
        .login = vLogin
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOUsuarioActivo, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .ip = pQuery.Fields(.ddl.ip)
        .hwnd = pQuery.Fields(.ddl.hwnd)
        .login = pQuery.Fields(.ddl.login)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.ip & " = '" & replaceSpecialSymbols(vIP) & "'"
    strSQL = strSQL & " AND " & ddl.hwnd & " = " & vHwnd

    makeWherePrimaryKey = strSQL

End Function

Private Sub findSQL(pSQL As String, pDB As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = pDB.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

Public Sub findByPrimaryKey(Optional pDB As Variant)

    findByPrimaryKey_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findByPrimaryKey_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.ip & " = '" & replaceSpecialSymbols(vIP) & "'"
    strSQL = strSQL & " ORDER BY " & ddl.hwnd & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Function delete(Optional pDB As Variant) As Boolean

    delete = delete_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function delete_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete_ = pDB.execute(strSQL)

End Function

Public Function save(Optional pDB As Variant) As Boolean

    save = save_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function save_(pDB As clsDB) As Boolean

    If exist_(pDB) Then
        save_ = update_(pDB)
    Else
        save_ = add_(pDB)
    End If

End Function

Public Function exist(Optional pDB As Variant) As Boolean

    exist = exist_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function exist_(pDB As clsDB) As Boolean
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)
    exist_ = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(Optional pDB As Variant) As Boolean

    add = add_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function add_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.login & " = '" & replaceSpecialSymbols(vLogin) & "'"
    strSQL = strSQL & ", " & ddl.ip & " = '" & replaceSpecialSymbols(vIP) & "'"
    strSQL = strSQL & ", " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findByPrimaryKey_ pDB

End Function

Public Function update(Optional pDB As Variant) As Boolean

    update = update_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function update_(pDB As clsDB) As Boolean
Dim strSQL As String
Dim strSET As String

Dim blnChange As Boolean

    update_ = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vLogin <> clon.login Or IsNull(vLogin) <> IsNull(clon.login) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.login & " = '" & replaceSpecialSymbols(vLogin) & "'"
        blnChange = True
    End If

    strSQL = "UPDATE " & ddl.table & " SET" & strSET
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    If blnChange Then update_ = pDB.execute(strSQL)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, Optional pDB As Variant)

    newID_ pFind, IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub newID_(pFind As Boolean, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    vHwnd = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.ip & " = '" & replaceSpecialSymbols(vIP) & "'"
    strSQL = strSQL & " ORDER BY " & ddl.hwnd & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vHwnd = 1 + rstQuery.Fields(ddl.hwnd)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOUsuarioActivo

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOUsuarioActivo
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Sub fillCombo(pCombo As ComboBox, Optional pDB As Variant)

    fillCombo_ pCombo, collectionAll(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOUsuarioActivo

    pCombo.Clear

    For Each objElemento In pCollection
        pCombo.AddItem objElemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = objElemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, Optional pDB As Variant)

    fillList_ pList, collectionAll(pDB)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim objElemento As clsDAOUsuarioActivo

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function hayUsuarioActivo(pIP As String, phWnd As Long, pDB As clsDB) As Boolean
    
    vIP = pIP
    vHwnd = phWnd
    
    hayUsuarioActivo = exist_(pDB)
    
End Function

Public Function eliminarUsuarioActivo(pIP As String, phWnd As Long, pDB As clsDB) As Boolean

    vIP = pIP
    vHwnd = phWnd
    
    eliminarUsuarioActivo = delete_(pDB)

End Function

Public Function grabaUsuario(pMDI As MDIForm, pUsuario As String, pDB As clsDB) As Boolean
    
    With Me
        .ip = pMDI.sokAplicacion.LocalIP
        .hwnd = pMDI.hwnd
        .login = pUsuario
        
        grabaUsuario = .save(pDB)
    End With

End Function


