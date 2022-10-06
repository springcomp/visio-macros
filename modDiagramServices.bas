Option Explicit

Private nDiagramServicesEnabled As Integer

Public Function EnableDiagramServices() As Long
    Let nDiagramServicesEnabled = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150
    Let EnableDiagramServices = nDiagramServicesEnabled
End Function

Public Function SetDiagramServices(ByVal nServicesEnabled)
    ActiveDocument.DiagramServicesEnabled = nServicesEnabled
End Function

Public Function ResetDiagramServices()
    SetDiagramServices nDiagramServicesEnabled
End Function
