Sub testeLU_vs1()
    Rows("1:12").Select
    Selection.Delete
    Range("A:A,F:G,I:K,P:Q,R:S,T:V").Select
    Selection.Delete Shift:=xlToLeft
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Pax"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "CPF/CNPJ"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "File"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Margem"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Recebimento"
    Columns("K:O").Select
    Selection.ColumnWidth = 15.29
    Columns("K:K").Select
    Selection.Copy
    Columns("L:O").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("a1").Select
End Sub

Sub testeLU_vs2()
    
    Dim lin As Integer
    Dim lin2 As Integer
    Dim varID As String
    Dim vAmount As Double
    Dim vFee As Double
    Dim vPayable As Double
        
        'copiar a base e colar numa nova aba
        Cells.Select
        Selection.Copy
        Sheets.Add After:=ActiveSheet
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        'apagar informaçoes desnecessarias
        Rows("1:12").Select
        Selection.Delete
        Range("A:A,F:G,I:K,P:Q,R:S,T:V").Select
        Selection.Delete Shift:=xlToLeft
        Range("K1").Select
        ActiveCell.FormulaR1C1 = "Pax"
        Range("L1").Select
        ActiveCell.FormulaR1C1 = "CPF/CNPJ"
        Range("M1").Select
        ActiveCell.FormulaR1C1 = "File"
        Range("N1").Select
        ActiveCell.FormulaR1C1 = "Margem"
        Range("O1").Select
        ActiveCell.FormulaR1C1 = "Recebimento"
        Columns("K:O").Select
        Selection.ColumnWidth = 10
        Columns("K:K").Select
        Selection.Copy
        Columns("a:O").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Range("a1").Select
            
        'Colocar as informações em apenas uma linha
        lin = 2
        lin2 = lin
        
        Do Until Cells(lin, "J") = Empty Or Cells(lin, "J") = ""
            varID = Cells(lin, "J")
            Do Until CStr(Cells(lin2, "j")) <> varID
                lin2 = lin2 + 1
            Loop
            
            vAmount = WorksheetFunction.SUM(Range(Cells(lin, "F"), Cells(lin2 - 1, "F")))
            vFee = WorksheetFunction.SUM(Range(Cells(lin, "G"), Cells(lin2 - 1, "G")))
            vPayable = WorksheetFunction.SUM(Range(Cells(lin, "H"), Cells(lin2 - 1, "H")))
            Range(Rows(lin + 1), Rows(lin2 - 1)).Select
            Selection.Delete
            
            Cells(lin, "F").Value = vAmount
            Cells(lin, "G").Value = vFee * -1
            Cells(lin, "H").Value = vPayable
            
            lin = lin + 1
            lin2 = lin
        Loop
        
End Sub

