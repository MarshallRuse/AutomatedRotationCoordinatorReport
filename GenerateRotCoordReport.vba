Option Explicit

Sub Generate_Rot_Coord_Block_Report()

    Dim DataWorkbook As Workbook
    
    Application.ScreenUpdating = False
    Call GenerateRotCoordExtract.Generate_Rotation_Coordinator_Extract(DataWorkbook)
    
    ' For some reason the screen was still updating unless I explicitly turned it on and then back
    ' off again between calling subs
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Call FormatBlockReport.Format_Block_Report(DataWorkbook)
    Application.ScreenUpdating = True

End Sub
