Sub Macro2()
'
' Macro2 Macro
'
' Atalho do teclado: Ctrl+Shift+Q
'
    Range("B6").Select
    Windows("PERSONAL.XLSB").Activate
    Range("B2:B3").Select
    Selection.Copy
    Windows("ADF CONTRA KOMPLETA JUL 19 -.xlsx").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("ND AD-V CONTRA AD-E MAI19").Select
    Sheets("ND AD-V CONTRA AD-E MAI19").Name = "ND AD-V CONTRA A D-E MAI19"
    Application.CutCopyMode = False
    ActiveWorkbook.Save
End Sub
