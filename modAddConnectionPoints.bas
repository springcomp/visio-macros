Option Explicit

Public Sub AddSnapPoints()
' Adds evenly distributed connection points to a rectangular shape
'
    
    Dim vShape As Visio.Shape
    
    'Requires a shape to be selected.
    Set vShape = GetSelectedShape()
    
    ' Add connection points evenly distributed
    ' around the rectangular shape
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0", _
        "Height*0"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.25", _
        "Height*0"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.5", _
        "Height*0"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.75", _
        "Height*0"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*1", _
        "Height*0"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*1", _
        "Height*0.25"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*1", _
        "Height*0.5"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*1", _
        "Height*0.75"
    
    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*1", _
        "Height*1"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.75", _
        "Height*1"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.5", _
        "Height*1"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0.25", _
        "Height*1"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0", _
        "Height*1"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0", _
        "Height*0.75"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0", _
        "Height*0.5"

    SetConnectionPoint _
        AddRowToConnectionPointsSection(vShape), _
        "Width*0", _
        "Height*0.25"

End Sub

Private Sub SetConnectionPoint(ByRef vRow As Visio.Row, sFormulaX As String, sFormulaY As String)

    With vRow
        .Cell(visCnnctX).FormulaU = sFormulaX
        .Cell(visCnnctY).FormulaU = sFormulaY
    
        .Cell(visCnnctDirX).FormulaU = 1#
        .Cell(visCnnctDirY).FormulaU = 0#
        .Cell(visCnnctType).FormulaU = visCnnctTypeInward
    End With

End Sub

Private Function AddRowToConnectionPointsSection(ByRef vShape As Visio.Shape) As Visio.Row

    Dim nRowIndex As Integer
    nRowIndex = vShape.AddRow(visSectionConnectionPts, visRowLast, visTagCnnctPt)
                    
    Dim vRow As Visio.Row
    Set vRow = GetConnectionPointsSection(vShape).Row(nRowIndex)

    Set AddRowToConnectionPointsSection = vRow
End Function

Private Function GetConnectionPointsSection(ByRef vShape As Visio.Shape) As Visio.Section
    Set GetConnectionPointsSection = vShape.Section(visSectionConnectionPts)
End Function

