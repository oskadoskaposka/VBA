Sub texto_para_data()
    
    Cells.Select
    Selection.UnMerge
    
    Columns("DA:DA").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("DA8").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-1])"
    Range("DB8").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-2])"
    Range("DC8").Select
    ActiveCell.FormulaR1C1 = "=DAY(RC[-3])"
    Range("DD8").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(DATE(RC[-3],RC[-2],RC[-1]),"""")"
    
    Range("DA8:DD8").Select
    Range("DD8").Activate
    
    'colocar aqui
    Range("de8").Select
    Selection.End(xlDown).Select
    lin = ActiveCell.Row
    col = ActiveCell.Column
    Range(Cells(lin, col - 4), Cells(lin, col - 1)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    Range("DD:DD").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("CZ:DC").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("CZ7").Value = "Vcts. Cliente"
    
    
End Sub

