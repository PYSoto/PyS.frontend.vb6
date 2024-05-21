Attribute VB_Name = "modGrid"
Option Explicit

Public Sub makeGrid(pGrid As MSFlexGrid, pTitulos As Variant, pAnchos As Variant, pColumnasFijas As Integer, pFilasFijas As Integer, pMode As Integer)
Dim intColumnas As Integer
Dim intColumna As Integer
    
    intColumnas = UBound(pTitulos) + 1
    pGrid.Clear
    pGrid.SelectionMode = pMode
    pGrid.ScrollBars = flexScrollBarVertical
    pGrid.Rows = 2
    pGrid.Cols = intColumnas
    pGrid.FixedCols = pColumnasFijas
    pGrid.FixedRows = pFilasFijas
    pGrid.Rows = 1
    For intColumna = 0 To intColumnas - 1
        pGrid.TextMatrix(0, intColumna) = pTitulos(intColumna)
        pGrid.ColWidth(intColumna) = pAnchos(intColumna)
    Next intColumna

End Sub

Public Sub makeGrid2(pGrid As MSFlexGrid, pTitulos As Variant, pColumnasFijas As Integer, pFilasFijas As Integer, pMode As Integer)
Dim intColumnas As Integer
Dim intColumna As Integer
    
    intColumnas = UBound(pTitulos) + 1
    pGrid.Clear
    pGrid.SelectionMode = pMode
    pGrid.ScrollBars = flexScrollBarVertical
    pGrid.Rows = 2
    pGrid.Cols = intColumnas
    pGrid.FixedCols = pColumnasFijas
    pGrid.FixedRows = pFilasFijas
    pGrid.Rows = 1
    For intColumna = 0 To intColumnas - 1
        pGrid.TextMatrix(0, intColumna) = pTitulos(intColumna)(0)
        pGrid.ColWidth(intColumna) = pTitulos(intColumna)(1)
    Next intColumna

End Sub

Public Sub makeGridCols(pGrid As MSFlexGrid, pTitulos As Variant, pAnchos As Variant, pFechaInicial As Date, pColumnasExtras As Integer, pMultiplicidad As Integer, pProporciones As Variant, pPosTitulo As Integer, pColumnasFijas As Integer, pFilasFijas As Integer, pMode As Integer)
Dim intColumnas As Integer
Dim intColumna As Integer
Dim intAncho As Integer
Dim intProporcion As Integer
Dim intCiclo As Integer
Dim intAnchoExtra As Integer
Dim intOffset As Integer
    
    intColumnas = UBound(pTitulos) + 1
    pGrid.Clear
    pGrid.SelectionMode = pMode
    pGrid.ScrollBars = flexScrollBarVertical
    pGrid.Rows = 2
    pGrid.Cols = intColumnas + pColumnasExtras * pMultiplicidad
    pGrid.FixedCols = pColumnasFijas
    pGrid.FixedRows = pFilasFijas
    pGrid.Rows = 1
    
    intAncho = 0
    For intColumna = 0 To intColumnas - 1
        pGrid.TextMatrix(0, intColumna) = pTitulos(intColumna)
        pGrid.ColWidth(intColumna) = pAnchos(intColumna)
        intAncho = intAncho + pAnchos(intColumna)
    Next intColumna
    
    intAnchoExtra = (pGrid.Width - 500 - intAncho) / pColumnasExtras
    
    intColumna = intColumnas
    For intCiclo = 1 To pColumnasExtras
        intOffset = 0
        For intProporcion = LBound(pProporciones) To UBound(pProporciones)
            If intProporcion = pPosTitulo - 1 Then pGrid.TextMatrix(0, intColumna) = Day(pFechaInicial + intCiclo - 1)
            intAncho = CInt(CDbl(intAnchoExtra) * CDbl(pProporciones(intProporcion)) / 100)
            pGrid.ColWidth(intColumna) = intAncho - intOffset
            intOffset = intOffset + pGrid.ColWidth(intColumna) - intAncho
            intColumna = intColumna + 1
        Next intProporcion
    Next intCiclo

End Sub

Public Sub setCheckGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer, pValor As Boolean)

    pGrid.Row = pRow
    pGrid.Col = pCol
    pGrid.CellFontName = "Wingdings"
    pGrid.CellFontSize = 14
    pGrid.CellAlignment = 4
    pGrid.Text = IIf(pValor, Chr(254), Chr(113))

End Sub

Public Function valueCheckGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer) As Boolean
    
    valueCheckGrid = IIf(Asc(pGrid.TextMatrix(pRow, pCol)) = 254, True, False)

End Function

Public Function topGrid(pGrid As MSFlexGrid) As Long
    
    topGrid = pGrid.Top + pGrid.CellTop

End Function

Public Function leftGrid(pGrid As MSFlexGrid) As Long
    
    leftGrid = pGrid.Left + pGrid.CellLeft

End Function

Public Sub setColourGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer, pColor As Variant)
    
    pGrid.Row = pRow
    pGrid.Col = pCol
    pGrid.CellBackColor = modColor.variant2RGB(pColor)

End Sub

Public Function getColourGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer) As Variant
    
    pGrid.Row = pRow
    pGrid.Col = pCol
    getColourGrid = pGrid.CellBackColor

End Function

Public Sub swapColor(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer, pColor1 As Variant, pColor2 As Variant)

    setColourGrid pGrid, pRow, pCol, IIf(getColourGrid(pGrid, pRow, pCol) = pColor1, pColor2, pColor1)
    
End Sub

Public Sub fieldAdd(pLine As String, pField As Variant)

    pLine = pLine & pField & Chr(9)
    
End Sub

Public Function array2itemGrid(pArray As Variant) As String
Dim strLinea As String

Dim intCiclo As Integer

    strLinea = ""
    
    For intCiclo = LBound(pArray) To UBound(pArray)
        fieldAdd strLinea, pArray(intCiclo)
    Next intCiclo
    
    array2itemGrid = strLinea
    
End Function

Public Sub setTextBox(pGrid As MSFlexGrid, pTextBox As TextBox)

    With pTextBox
        .Visible = False
        .Left = leftGrid(pGrid)
        .Top = topGrid(pGrid)
        .Width = pGrid.CellWidth
        .Height = pGrid.CellHeight
        .Visible = True
        .Text = pGrid.Text
        .SetFocus
    End With
    
End Sub

Public Sub makeGridColsHoras(pGrid As MSFlexGrid, pTitulos As Variant, pAnchos As Variant, pInicio As Integer, pColumnasExtras As Integer, pMultiplicidad As Integer, pProporciones As Variant, pPosTitulo As Integer, pColumnasFijas As Integer, pFilasFijas As Integer, pMode As Integer)
Dim intColumnas As Integer
Dim intColumna As Integer
Dim intAncho As Integer
Dim intProporcion As Integer
Dim intCiclo As Integer
Dim intAnchoExtra As Integer
Dim intOffset As Integer
    
    intColumnas = UBound(pTitulos) + 1
    pGrid.Clear
    pGrid.SelectionMode = pMode
    pGrid.ScrollBars = flexScrollBarVertical
    pGrid.Rows = 2
    pGrid.Cols = intColumnas + pColumnasExtras * pMultiplicidad
    pGrid.FixedCols = pColumnasFijas
    pGrid.FixedRows = pFilasFijas
    pGrid.Rows = 1
    
    intAncho = 0
    For intColumna = 0 To intColumnas - 1
        pGrid.TextMatrix(0, intColumna) = pTitulos(intColumna)
        pGrid.ColWidth(intColumna) = pAnchos(intColumna)
        intAncho = intAncho + pAnchos(intColumna)
    Next intColumna
    
    intAnchoExtra = (pGrid.Width - 500 - intAncho) / pColumnasExtras
    
    intColumna = intColumnas
    For intCiclo = 1 To pColumnasExtras
        intOffset = 0
        For intProporcion = LBound(pProporciones) To UBound(pProporciones)
            If intProporcion = pPosTitulo - 1 Then pGrid.TextMatrix(0, intColumna) = (pInicio + intCiclo - 1)
            intAncho = CInt(CDbl(intAnchoExtra) * CDbl(pProporciones(intProporcion)) / 100)
            pGrid.ColWidth(intColumna) = intAncho - intOffset
            intOffset = intOffset + pGrid.ColWidth(intColumna) - intAncho
            intColumna = intColumna + 1
        Next intProporcion
    Next intCiclo

End Sub


