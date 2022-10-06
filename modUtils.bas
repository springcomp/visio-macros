Option Explicit

Public Function GetSelectedShape() As Visio.Shape
    
    Dim vShape As Visio.Shape
    
    If Visio.ActiveWindow.Selection.Count <> 1 Then
        MsgBox "Select a shape, then re-run macro."
        Exit Function
    Else
        Set vShape = ActiveWindow.Selection(1)
    End If
    
    Set GetSelectedShape = vShape

End Function

