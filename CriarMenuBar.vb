Option Explicit
    
    'definir os nomes das barras
    Public Const ToolBarName   As String = "Macros úteis"
    Public Const ToolBarName2   As String = "Macros úteis2"

Sub Auto_Open()
    
    'rodar a macro ao abrir a planilha
    Call CreateMenubar

End Sub

Sub Auto_Close()
    
    'ao fechar a planilha remover a barra
    Call RemoveMenubar

End Sub

Sub RemoveMenubar()
  
  On Error Resume Next
  
  'apagar as barras
  Application.CommandBars(ToolBarName).Delete
  Application.CommandBars(ToolBarName2).Delete
  
  On Error GoTo 0
  
End Sub

Sub CreateMenubar()

    'definir as variaveis
    Dim iCtr As Long
    Dim MacNames As Variant
    Dim MacNames2 As Variant
    Dim CapNamess As Variant
    Dim CapNamess2 As Variant
    Dim TipText As Variant
    Dim TipText2 As Variant
        
        'apagar as barras antigas
        Call RemoveMenubar
        
        'definir as macros que vão para a barra1
        MacNames = Array("Esconder_aba", _
                      "Mostrar_abas_ocultas", "Quebrar_Senha", "Desproteger_Todas", _
                      "Proteger_Todas", "Desproteger_csenhapadraoPaty", _
                      "Proteger_csenhapadraoPaty", "Copiar_ColarValores", _
                      "Alterar_Propriedades_das_imagens_para_Moveandsize", "Nao_alterar_imagens_com_celulas", _
                      "estilo_a1", "estilo_l1c1", "limpar_tabela")
        
        'definir o titulo que aparecerá na barra1
        CapNamess = Array("Esconder aba", _
                      "Mostrar todas", "Quebrar senha", "Desproteger todas", _
                      "Proteger todas", "Desproteger aba", _
                      "Proteger aba", "Colar valor", _
                      "Alterar imagens", "Não alterar imagens", _
                      "Estilo A1", "Estilo L1C1", "Limpar Tabela")
        
        'definir o texto que aparecerá ao deixar o mouse parado na barra1
        TipText = Array("Esconder aba.", _
                    "Mostrar todas abas da planilha.", " Quebrar a senha da aba.", "Desproteger todas as abas de uma vez.", _
                    "Proteger todas as abas de uma vez.", "Desproteger a aba com a senha padrão Paty.", _
                    "Proteger a aba com a senha padrão Governança.", "Copiar o conteúdo da seleção e colar como valores.", _
                    "Alterar imagens com células", "Não alterar as imagens com células", _
                    "Difinir estilo A1", "Definir estilo L1C1", "123")
        
        '
        With Application.CommandBars.Add
          .Name = ToolBarName
          .Left = 200
          .Top = 200
          .Protection = msoBarNoProtection
          .Visible = True
          .Position = msoBarFloating
        
          For iCtr = LBound(MacNames) To UBound(MacNames)
            With .Controls.Add(Type:=msoControlButton)
              .OnAction = "'" & ThisWorkbook.Name & "'!" & MacNames(iCtr)
              .Caption = CapNamess(iCtr)
              .Style = msoButtonIconAndCaption
              .FaceId = 71 + iCtr
              .TooltipText = TipText(iCtr)
            End With
          Next iCtr
          
        End With
        
    
    
    
End Sub


