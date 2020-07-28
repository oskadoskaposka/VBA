Sub COPIAR_COLAR_VALORES()
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

'Função para comparar se a variavel esta em uma lista
    'exEMPLO:
    'If IsIn(lin, "8,9,10,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36") Then Call double_click_no_fluxo

Public Function IsIn(valCheck, valList As String) As Boolean

    IsIn = Not InStr("," & valList & ",", "," & valCheck & ",") = 0
    
End Function

Sub LIMPAR_FILTRO_EMPRESA()
    
    Range("F801").Select
    ActiveSheet.ShowAllData

End Sub

Sub FILTRAR_EMPRESA()
    
    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14, Criteria1:="AD"
    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14, Criteria1:="DIVERSA SANTANDER"

End Sub

Sub FILTRAR_DATA()
    
    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8, Operator:=xlFilterValues, Criteria2:=Array(2, "1/2/2019")
    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8, Operator:=xlFilterValues, Criteria2:=Array(2, "1/2/2019", 2, "1/7/2019", 2, "1/8/2019", 2, "1/9/2019", 2, "1/10/2019", 2, "1/11/2019", 2, "1/14/2019", 2, "1/15/2019", 2, "1/16/2019", 2, "1/21/2019", 2, "1/23/2019", 2, "1/24/2019", 2, "1/25/2019", 2, "1/29/2019")

End Sub

'Mudar planilha entre modo de calculo manual e automatico
Sub MAN_AUTO()

    If Application.Calculation = xlManual Then
        Application.Calculation = xlAutomatic
        MsgBox ("PLANILHA EM CALCULO AUTOMATICO")
    ElseIf Application.Calculation = xlAutomatic Then
        Application.Calculation = xlManual
        MsgBox ("PLANILHA EM CALCULO MANUAL")
    End If
   
End Sub

'Mudar planilha entre modo getpivotdata (ao clicar na tabela dinamica montar formula de tabela dinamica) e referencia normal
Sub GET_PIVOT_DATA()

    If Application.GenerateGetPivotData = True Then
        Application.GenerateGetPivotData = False
    ElseIf Application.GenerateGetPivotData = False Then
        Application.GenerateGetPivotData = True
    End If

    Application.GenerateTableRefs = xlGenerateTableRefStruct
    
End Sub


'Limpar todas informações de uma tabela, mantem cabeçalho
Sub limpar_tabela()
 
    Dim ws As Worksheet
    Dim obj As ListObject
    Dim ListR As ListRows

    Set ws = ActiveSheet
    Set obj = ws.ListObjects(1)
    Set ListR = obj.ListRows
   
    With obj
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With

End Sub

'Função para somar/contar as celulas de acordo com a cor usada nela
'para usar a fómula é
'=colorfunction("Celula com a cor";"range para somar/contar";"Falso=contar, Verdadeiro=somar")

'Definindo como a formula funciona
Function ColorFunction(rColor As Range, rRange As Range, Optional SUM As Boolean)
    
    Dim rCell As Range
    Dim lCol As Long
    Dim vResult
    
        lCol = rColor.Interior.ColorIndex

        'Se for soma, some
        If SUM = True Then
            For Each rCell In rRange
                If rCell.Interior.ColorIndex = lCol Then
                    vResult = WorksheetFunction.SUM(rCell, vResult)
                End If
            Next rCell
        'Se não for soma, conte
        Else
            For Each rCell In rRange
                If rCell.Interior.ColorIndex = lCol Then
                    vResult = 1 + vResult
                End If
            Next rCell
        End If
        
        'Resultado da fórmula
        ColorFunction = vResult

End Function
