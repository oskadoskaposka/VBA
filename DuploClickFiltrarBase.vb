


'COLOCAR ABAIXO NA PLANILHA ESCOLHIDA
'Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    'Desabilitar a edição no duplo clique no RANGE SELECIONADO
    'If Not Intersect(Target, Range("a6:gk73")) Is Nothing Then Cancel = True
    
    'RODAR A MACRO
    'Call Aplicar_Filtros

'End Sub

'Function IsIn(valCheck, valList As String) As Boolean
'    IsIn = Not InStr("," & valList & ",", "," & valCheck & ",") = 0
'End Function

'Sub double_click_no_fluxo()
'
'    'Definir variaveis do endereço
'    Dim empreSA As String
'    Dim DAta As String
'    Dim coNTa As String
'
'        'Otimizar a velocidade da macro
'        Application.ScreenUpdating = False
'        Application.Calculation = xlCalculationManual
'        Application.EnableEvents = False'
'
 '       'Definir a empresa, data e conta de acordo com o duplo click
  '      'endEReco = ActiveCell.Address
   '     empreSA = Fluxo.Cells(1, "A").Value
    '    DAta = Cells(6, ActiveCell.Column)
     '   coNTa = Cells(ActiveCell.Row, "A")
      '
       ' 'Limpar filtro da base e reaplicar com as informações coletadas acima
        'bFluxo.Select
'        ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=1
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=2
  '      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=3
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=4
    '    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=5
     '   ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=6
      '  ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=7
       ' ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8
        'ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=9
'        ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=10
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=11
  '      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=12
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=13
    '    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14
     '   ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=15
        
        'empresa
'        ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14, Criteria1:=empreSA
        'data
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8, Criteria1:=DAta
        'conta
  '      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=6, Criteria1:=coNTa
        
        'Mostrar dados tabela
   '     bFluxo.Cells(4, "b").Activate
        
        'Habilitar atualização tela
    '    Application.EnableEvents = True
     '   Application.Calculation = xlCalculationAutomatic
      '  Application.ScreenUpdating = True
                
'End Sub

'Sub double_click_na_data()
 ''
   ' 'Definir variaveis do endereço
    'Dim empreSA As String
    'D im DAta As String
    
'        'Otimizar a velocidade da macro
 '       Application.ScreenUpdating = False
  '      Application.Calculation = xlCalculationManual
   '     Application.EnableEvents = False
'
        'Definir a empresa, data e conta de acordo com o duplo click
 '       empreSA = Fluxo.Cells(1, "A").Value
  '      DAta = Cells(6, ActiveCell.Column)
        
        'Limpar filtro da base e reaplicar com as informações coletadas acima
'        bFluxo.Select
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=1
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=2
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=3
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=4
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=5
 ''       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=6
  '      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=7
  ''      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=9
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=10
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=11
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=12
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=13
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=15
        
        'empresa
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14, Criteria1:=empreSA
        'data
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8, Criteria1:=DAta
        
        'Mostrar dados tabela
   '     bFluxo.Cells(4, "b").Activate
        
        'Habilitar atualização tela
    '    Application.EnableEvents = True
     '   Application.Calculation = xlCalculationAutomatic
      '  Application.ScreenUpdating = True
                
'End Sub

'Sub double_click_na_conta()
    
    'Definir variaveis do endereço
 '   Dim empreSA As String
  '  Dim coNTa As String
    
        'Otimizar a velocidade da macro
   '     Application.ScreenUpdating = False
    '    Application.Calculation = xlCalculationManual
     '   Application.EnableEvents = False

        'Definir a empresa, data e conta de acordo com o duplo click
'        empreSA = Fluxo.Cells(1, "A").Value
 '       coNTa = Cells(ActiveCell.Row, "A")
        
        'Limpar filtro da base e reaplicar com as informações coletadas acima
  '      bFluxo.Select
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=1
    '    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=2
     '   ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=3
      '  ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=4
       ' ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=5
        'ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=6
'        ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=7
 '       ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=8
  '      ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=9
   '     ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=10
    '    ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=11
     '   ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=12
      '  ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=13
       ' ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14
        'ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=15
        
        'empresa
       ' ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=14, Criteria1:=empreSA
        'conta
       ' ActiveSheet.ListObjects("LCTOS").Range.AutoFilter Field:=6, Criteria1:=coNTa
        
        'Mostrar dados tabela
       ' bFluxo.Cells(4, "b").Activate
        
        'Habilitar atualização tela
       ' Application.EnableEvents = True
       ' Application.Calculation = xlCalculationAutomatic
        'Application.ScreenUpdating = True
                
'End Sub

'Sub Aplicar_Filtros()
'
  '  Application.ScreenUpdating = False
    
 '   Fluxo.Select

   ' Dim col As String
    'Dim lin As String
    
    'fazer a macro rodar apenas nas linhas com informação do fluxo
'    col = ActiveCell.Column
 '   lin = ActiveCell.Row
               
    'Liberar a macro nas linhas de emprestimo
    'jan
'    If IsIn(lin, "8,9,10") Then If IsIn(col, "3,5,7,9,11,13,15,17,19,21,23,25,27,29,31,33,35,37,39,41,43,45,47,49,51,53,55,57,59,61,63") Then Call double_click_no_fluxo
    'fev
'    If IsIn(lin, "8,9,10") Then If IsIn(col, "70,72,74,76,78,80,82,84,86,88,90,92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122,124") Then Call double_click_no_fluxo
    'mar
  '  If IsIn(lin, "8,9,10") Then If IsIn(col, "131,133,135,137,139,141,143,145,147,149,151,153,155,157,159,161,163,165,167,169,171,173,175,177,179,181,183,185,187,189,191") Then Call double_click_no_fluxo
    
    'Liberar a macro nas linhas de receita
   ' If IsIn(lin, "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36") Then If IsIn(col, "3,5,7,9,11,13,15,17,19,21,23,25,27,29,31,33,35,37,39,41,43,45,47,49,51,53,55,57,59,61,63") Then Call double_click_no_fluxo
    'fev
    'If IsIn(lin, "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36") Then If IsIn(col, "70,72,74,76,78,80,82,84,86,88,90,92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122,124") Then Call double_click_no_fluxo
    'mar
'    If IsIn(lin, "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36") Then If IsIn(col, "131,133,135,137,139,141,143,145,147,149,151,153,155,157,159,161,163,165,167,169,171,173,175,177,179,181,183,185,187,189,191") Then Call double_click_no_fluxo
    
    'Liberar a macro nas linhas de despesa
 '   If IsIn(lin, "40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63") Then If IsIn(col, "3,5,7,9,11,13,15,17,19,21,23,25,27,29,31,33,35,37,39,41,43,45,47,49,51,53,55,57,59,61,63") Then Call double_click_no_fluxo
    'fev
  '  If IsIn(lin, "40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63") Then If IsIn(col, "70,72,74,76,78,80,82,84,86,88,90,92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122,124") Then Call double_click_no_fluxo
    'mar
   ' If IsIn(lin, "40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63") Then If IsIn(col, "131,133,135,137,139,141,143,145,147,149,151,153,155,157,159,161,163,165,167,169,171,173,175,177,179,181,183,185,187,189,191") Then Call double_click_no_fluxo

    'Liberar a macro nas linhas de data
'    If IsIn(lin, "6") Then If IsIn(col, "3,5,7,9,11,13,15,17,19,21,23,25,27,29,31,33,35,37,39,41,43,45,47,49,51,53,55,57,59,61,63,70,72,74,76,78,80,82,84,86,88,90,92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122,124,131,133,135,137,139,141,143,145,147,149,151,153,155,157,159,161,163,165,167,169,171,173,175,177,179,181,183,185,187,189,191") Then Call double_click_na_data

    'Liberar a macro nas linhas de conta
 '   If IsIn(lin, "8,9,10,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63") Then If IsIn(col, "1") Then Call double_click_na_conta

 '   Application.ScreenUpdating = True
'
'End Sub


