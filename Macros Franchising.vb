Sub SemFilial()

Application.ScreenUpdating = False
    Windows("PERSONAL.XLSB").Visible = True
    Range("O4:S34").Select
    Selection.Copy
    ActiveWindow.Visible = False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
Application.ScreenUpdating = True
    
End Sub

Sub ComFilial()

Application.ScreenUpdating = False
    Windows("PERSONAL.XLSB").Visible = True
    Range("O38:S68").Select
    Selection.Copy
    ActiveWindow.Visible = False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
Application.ScreenUpdating = True
    
End Sub

Sub Aud_4()

Application.ScreenUpdating = False
    Windows("PERSONAL.XLSB").Visible = True
    Range("O78:T88").Select
    Selection.Copy
    ActiveWindow.Visible = False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
Application.ScreenUpdating = True
    
End Sub

