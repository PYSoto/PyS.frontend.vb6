VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public xlsServicios As Object
Public xlsWorkBook As Object
Public xlsActiveSheet As Object

Public Sub fillCell(pRow As Integer, pCol As Integer, pBold As Boolean, pValor As String)
    
    With xlsActiveSheet
        .cells(pRow, pCol).Font.Bold = pBold
        .cells(pRow, pCol).Value = pValor
    End With
    
End Sub

Public Sub fillCellNumber(pRow As Integer, pCol As Integer, pBold As Boolean, ByVal pValor As Double)
    
    With xlsActiveSheet
        .cells(pRow, pCol).Font.Bold = pBold
        .cells(pRow, pCol).Value = pValor
    End With

End Sub

Public Sub newExcel()

    Set xlsServicios = CreateObject("Excel.Application")
    Set xlsWorkBook = xlsServicios.Workbooks.add
    xlsServicios.Visible = False
    Set xlsActiveSheet = xlsWorkBook.Worksheets(1)

End Sub

Public Sub openExcel(pArchivo As String)

    Set xlsServicios = CreateObject("Excel.Application")
    Set xlsWorkBook = xlsServicios.Workbooks.Open(pArchivo)
    xlsServicios.Visible = False
    Set xlsActiveSheet = xlsWorkBook.Worksheets(1)

End Sub

Public Sub saveExcel(pArchivo As String)

On Error Resume Next
    
    Kill pArchivo
    
    xlsWorkBook.SaveAs pArchivo
    Set xlsActiveSheet = Nothing
    xlsServicios.Quit
    Set xlsServicios = Nothing
       
End Sub

