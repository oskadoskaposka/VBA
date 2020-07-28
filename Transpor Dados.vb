Sub Transposição_BD_Recebível()

    Sheets("BD_COM_JUROS_new").Select
    Range("BD_COM_JUROS[[#Headers],[chave]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("antigo_BD_COM_JUROS").Select
    Range("BD_COM_JUROS_DATA[[#Headers],[chave]]").Select
    ActiveCell.Offset(1, 1).Select
    Call limpar_tabela
    Sheets("BD_COM_JUROS_new").Select
    Range("BD_COM_JUROS[[#Headers],[chave]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("antigo_BD_COM_JUROS").Select
    Range("BD_COM_JUROS_DATA[[#Headers],[chave]]").Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.ListObject.ListRows(1).Delete
    Sheets("BD_COM_JUROS_new").Select
    Range("BD_COM_JUROS[[#Headers],[chave]]").Select
    ActiveCell.Offset(1, 1).Select
    Call limpar_tabela

End Sub
