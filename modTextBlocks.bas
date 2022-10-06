Public Sub MakeBackgroundTextTransparent()

    Dim vShape As Visio.Shape
    
    'Requires a shape to be selected.
    Set vShape = GetSelectedShape()

    'Enable diagram services
    Dim nServicesEnabled As Long
    Let nServicesEnabled = EnableDiagramServices

    'Make background text transparent
    vShape.CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaU = "100%"

    'Restore diagram services
    SetDiagramServices nServicesEnabled

End Sub
