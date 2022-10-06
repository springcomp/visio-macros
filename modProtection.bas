Option Explicit

Public Sub ProtectShapeGroup()
'
' Protects a group so as to prevent member selections.
' - Locks aspect ratio and theme formatting
' - Disable member selection
'
    
    'Enable diagram services
    Dim nServicesEnabled As Long
    Let nServicesEnabled = EnableDiagramServices

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Propriétés de protection")
    
    Dim vShape As Visio.Shape
    Set vShape = GetSelectedShape()
    With vShape
        .CellsSRC(visSectionObject, visRowLock, visLockAspect).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockFromGroupFormat).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockThemeColors).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockThemeEffects).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockThemeConnectors).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockThemeFonts).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowLock, visLockThemeIndex).FormulaU = "1"
        .CellsSRC(visSectionObject, visRowThemeProperties, visColorSchemeIndex).FormulaU = "0"
        .CellsSRC(visSectionObject, visRowThemeProperties, visEffectSchemeIndex).FormulaU = "0"
        .CellsSRC(visSectionObject, visRowThemeProperties, visConnectorSchemeIndex).FormulaU = "0"
        .CellsSRC(visSectionObject, visRowThemeProperties, visFontSchemeIndex).FormulaU = "0"
        .CellsSRC(visSectionObject, visRowThemeProperties, visThemeIndex).FormulaU = "0"
    End With
    
    'With vShape
    '    .CellsSRC(visSectionObject, visRowGroup, visGroupSelectMode).FormulaU = "0"
    'End With
    
    Application.EndUndoScope UndoScopeID1, True

    'Restore diagram services
    SetDiagramServices nServicesEnabled

End Sub
