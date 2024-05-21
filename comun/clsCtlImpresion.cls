VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub printReport(pCrystalReport As CrystalReport, pReportName As String, pStringConnection As String, Optional pSubReports As Variant, Optional pParameters As Variant, Optional pSorts As Variant)
Dim varSubreport As Variant
Dim varParameter As Variant
Dim varSort As Variant

Dim blnExistSubReport As Boolean

Dim intSecuencia As Integer

Dim strParametros As String

On Error GoTo handler

    strParametros = ""
    
    If Not IsMissing(pParameters) Then
        For Each varParameter In pParameters
            strParametros = strParametros & "(" & varParameter(0) & "=" & varParameter(1) & ")"
        Next
    End If

    blnExistSubReport = False
    
    With pCrystalReport
        .Reset
        .ReportFileName = App.Path + "\..\Comun\" & pReportName & ".rpt"
        .Connect = pStringConnection
        
        If Not IsMissing(pSubReports) Then
            For Each varSubreport In pSubReports
                blnExistSubReport = True
                .SubreportToChange = varSubreport
                .Connect = pStringConnection
            Next
        End If
        
        If blnExistSubReport Then .SubreportToChange = ""
        
        If Not IsMissing(pParameters) Then
            intSecuencia = 0
            For Each varParameter In pParameters
                .ParameterFields(intSecuencia) = varParameter(0) & ";" & varParameter(1) & ";TRUE"
                intSecuencia = intSecuencia + 1
            Next
        End If
        
        If Not IsMissing(pSorts) Then
            intSecuencia = 0
            For Each varSort In pSorts
                .SortFields(intSecuencia) = varSort
                intSecuencia = intSecuencia + 1
            Next
        End If

        frmImpresora.Show vbModal
        frmImpresora.cargar pCrystalReport
        If Not frmImpresora.Cancel Then .Action = 1
    End With
    
    Exit Sub
    
handler:
    MsgBox "Error de Impresión: " & Err.Number & " - " & Err.Description
    Exit Sub
    
End Sub

