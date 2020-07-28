Option Explicit

'Desproteger Planilha
Sub Quebrar_Senha()

     Application.ScreenUpdating = False
 
        Dim i, i1, i2, i3, i4, i5, i6 As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer
        On Error Resume Next
        For i = 65 To 66
        For j = 65 To 66
        For k = 65 To 66
        For l = 65 To 66
        For m = 65 To 66
        For i1 = 65 To 66
        For i2 = 65 To 66
        For i3 = 65 To 66
        For i4 = 65 To 66
        For i5 = 65 To 66
        For i6 = 65 To 66
        For n = 32 To 126
        ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        If ActiveSheet.ProtectContents = False Then
        MsgBox "One usable password is " & Chr(i) & Chr(j) & _
        Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
        Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        Exit Sub
        End If
        Next
        Next
        Next
        Next
        Next
        Next
        Next
        Next
        Next
        Next
        Next
        Next
     
     Application.ScreenUpdating = True
     
End Sub

'Função que desprotege todas as planilhas de um arquivo
Sub Desproteger_Todas()

     Application.ScreenUpdating = False

'Declara as variáveis necessárias
Dim lPass As String
Dim lQtdePlan As Integer
Dim lPlanAtual As Integer
Dim lPlanOrigem As String
'Solicita a senha
'O método InputBox é utilizado para solicitar um valor através de um formulário
lPass = InputBox("Desproteger todas as planilhas:", "Senha")
'Inicia as variáveis
'O método Worksheets.Count passa a quantidade de planilhas existentes no arquivo
lQtdePlan = Worksheets.Count
lPlanAtual = 1
lPlanOrigem = ActiveSheet.Name
'Loop pelas planilhas
'A função While realiza um loop de código enquanto não passar por todas as planilhas contadas
While lPlanAtual <= lQtdePlan
'O método Worksheets(lPlanAtual).Activate ativa a planilha conforme o índice atual 1, 2, 3…
Worksheets(lPlanAtual).Activate
'O método .UnProtect desprotege a planilha
ActiveSheet.Unprotect Password:=lPass
'Muda o índice para passar para a próxima planilha
lPlanAtual = lPlanAtual + 1
Wend
Worksheets(lPlanOrigem).Activate
'O método MsgBox exibe um formulário de aviso ao usuário.
MsgBox "Planilhas desprotegidas!"

     Application.ScreenUpdating = True

End Sub

'Função que protege todas as planilhas de um arquivo
Sub Proteger_Todas()
'Declara as variáveis necessárias

Application.ScreenUpdating = False

Dim lPass As String
Dim lQtdePlan As Integer
Dim lPlanAtual As Integer
Dim lPlanOrigem As String
'Solicita a senha
'O método InputBox é utilizado para solicitar um valor através de um formulário
lPass = InputBox("Proteger todas as planilhas:", "Senha")
'Inicia As variáveis
'O método Worksheets.Count passa a quantidade de planilhas existentes no arquivo
lQtdePlan = Worksheets.Count
lPlanAtual = 1
lPlanOrigem = ActiveSheet.Name
'Loop pelas planilhas
'A função While realiza um loop de código enquanto não passar por todas as planilhas contadas
While lPlanAtual <= lQtdePlan
'O método Worksheets(lPlanAtual).Activate ativa a planilha conforme o índice atual 1, 2, 3…
Worksheets(lPlanAtual).Activate
'O método .Protect proteje a planilha passando os parâmetros para proteger
'objetos de desenho, conteúdo, cenários e passando o password digitado
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=lPass
'Muda o índice para passar para a próxima planilha
lPlanAtual = lPlanAtual + 1
Wend
Worksheets(lPlanOrigem).Activate
'O método MsgBox exibe um formulário de aviso ao usuário.
MsgBox "Planilhas protegidas!"

Application.ScreenUpdating = True

End Sub

'Retirar a senha padrão paty
Sub Desproteger_csenhapadraoPaty()

Application.ScreenUpdating = False

    ActiveSheet.Unprotect "8908"
    
Application.ScreenUpdating = True

End Sub

'Colocar a senha padrão paty
Sub Proteger_csenhapadraoPaty()

Application.ScreenUpdating = False

    ActiveSheet.Protect "8908"
    
Application.ScreenUpdating = True

End Sub

'Private Sub MacroErro()

'Application.ScreenUpdating = False

'Application.OnKey "^%{UP}", "Seta_Up"
'Application.OnKey "^%{DOWN}", "Seta_Down"

'Application.ScreenUpdating = True

'End Sub

'mostrar abas ocultas


Sub Mostrar_abas_vryhidden()

Application.ScreenUpdating = False

     Dim abas As Worksheet
     
     For Each abas In Worksheets
        If abas.Visible = xlSheetVeryHidden Then abas.Visible = xlSheetVisible
     Next abas
     
Application.ScreenUpdating = True

End Sub


Sub Mostrar_abas_ocultas()

Application.ScreenUpdating = False

     Dim abas As Worksheet
     
     For Each abas In Worksheets
        abas.Visible = xlSheetVisible
     Next abas
     
Application.ScreenUpdating = True

End Sub

Sub Esconder_aba()
On Error GoTo banana

Application.ScreenUpdating = False

    ActiveSheet.Visible = xlVeryHidden

banana:

Application.ScreenUpdating = True

End Sub

Sub Alterar_Propriedades_das_imagens_para_MoveAndSize()
'Sub Alterar_imagens_com_celulas ()

Application.ScreenUpdating = False

     ActiveSheet.DrawingObjects.Select
     Selection.Placement = xlMoveAndSize
     
Application.ScreenUpdating = True

End Sub

'Sub Nao_alterar_Propriedades_das_imagens_para_FreeFloating()
Sub Nao_alterar_imagens_com_celulas()

Application.ScreenUpdating = False

     ActiveSheet.DrawingObjects.Select
     Selection.Placement = xlFreeFloating
     
Application.ScreenUpdating = True

End Sub

Sub Copiar_ColarValores()

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
End Sub

Sub Apagar_e_inserir_coluna_A()

Application.ScreenUpdating = False

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToLeft
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = False
    Columns("A:A").Select
    
Application.ScreenUpdating = True
    
End Sub

Sub estilo_A1()
    
    Application.ReferenceStyle = xlA1

End Sub

Sub estilo_L1C1()
    
    Application.ReferenceStyle = xlR1C1

End Sub

Sub Colar_Valores()

On Error GoTo fim
    'Selection.PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
fim:

End Sub
