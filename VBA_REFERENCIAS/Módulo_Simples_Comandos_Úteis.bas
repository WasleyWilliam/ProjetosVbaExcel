Attribute VB_Name = "M�dulo_Simples_Comandos_�teis"
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRI��O.....:C�DIGOS SIMPLES
                                          'REFERENCIA....:
'==================================================================================================================
'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       - Simples C�digos para utiliza��o em projetos, c�digos variados
'
'PARA ATIVAR TECLA NUM DO TECLADO
'COLOCAR TEMPO ENTRE COMANDOS
'POPUP DE TEXTO
'POPUP QUE RECEBE INFORMA��O
'COMANDOS DE ALT+TAB E COMANDO TAB
'SELECIONAR UMA C�LULA/ ABA NA PLANILHA
'C�DIGO REPETI��O DO UNTIL (AT� A CONDI��O DE C�LULA VAZIA EM LINHA, COLUNA)
'ATIVAR E DESATIVAR MOVIMENTO DE TELA
'ABRIR E FECHAR UM FORMUL�RIO
'COMANDOS PARA MOVIMENTAR COM SETA TECLADO
'ATIVAR E DESATIVAR ABAS DE UMA PLANILHA / FORMUL�RIOS E MENUS DO EXCEL
'DESLOCAR X C�LULAS PARA O LADO COLULA, PARA LADOS DAS LINHAS
'CONDI��O IF (SE)
'PREENCHER CAIXAS DE COMBINA��O COM TEXTOS (DEVER� CRIAR VALIDA��O DE DADOS NAS C�LULAS ONDE CONSTAM OS VALORES.)
'ESCONDER PLANILHA � ABRIR SOMENTE FORMS
'ENCONTRAR VALOR DO N�MERO DA �LTIMA LINHA VAZIA
'LIMPAR C�LULAS DE ACORDO COM VALORES ISERIDOS
'ATIVAR E DESATIVAR ALERTAS DO EXCEL, TEMOS COMO EXEMPLO PERGUNTANDO SE DESEJA REALMENTE SAIR DE UM ARQUIVO SEM SALVAR
'TECLA DE ENTER / ESPA�O
'COLOCAR DATA NO FORMATO CORRETO (BRASIL)
'MODELOS DE DECLARA��O DE VARI�VEL
'BUSCAR NOME DO COMPUTADOR
'RETIRAR FILTRO DA PLANILHA
'INCLUIR FILTRO NA PLANILHA
'LIMPARA DADOS DA PLANILHA ( *LIMPAR� TODAS AS C�LULAS USADAS  DA PLANILHA ESPEC�FICA)
'SELECIONA E COPIA TODO ESPA�O UTILIZADO NA PLANILHA (*INCLUSIVE C�LULAS EM BRANCO)
'AJUSTAR TAMANHO DA COLUNA AUTOMATICAMENTE / CENTRALIZAR DADOS DE UMA C�LULA
'SALVAR PLANILHA
'COLOCAR ID PLANILHA
'LOCALIZAR OUTRO VALOR NA PLANILHA DE ACORDO COM UM VALOR DENTRO DE UMA TEXBOX
'SE UMA TECLA FOR PRESSIONADA. - ATEN��O CADA N�MERO EQUIVALEM A UMA TECLA, DEVEMOS CONSULTAR TABELAS.
'
'==================================================================================================================

'PARA ATIVAR TECLA NUM DO TECLADO
Sub ativar_numlook()
    Application.SendKeys "{NUMLOCK}", True
End Sub

'COLOCAR TEMPO ENTRE COMANDOS
Sub TEMPO()
    Application.Wait (Now + TimeValue("0:00:01"))
End Sub

'POPUP DE TEXTO
Sub Popup_msgbox()
    MSGBOX " DIGITE TEXTO DENTRO DAS ASPAS  "
End Sub

'POPUP QUE RECEBE INFORMA��O
Sub Popup_Inputbox()
    InputBox " SUA PERGUNTA AQUI  "
End Sub

'COMANDOS DE ALT+TAB E COMANDO TAB
Sub Alt_Tab()
    Application.SendKeys "%{Tab}"
    Application.SendKeys "{Tab}"
End Sub

'SELECIONAR UMA C�LULA/ ABA NA PLANILHA
Sub Select_Celula()
    Sheets("NOME DA PLANILHA").Select
    Range("c1").Select
    Cells(1, 1).Select
End Sub

'C�DIGO REPETI��O DO UNTIL (AT� A CONDI��O DE C�LULA VAZIA EM LINHA, COLUNA)
Sub DoUntil()
    Do Until Cells(LINHA, 1) = ""
'*Inserir C�digo entre Do Until e Loop � c�digo ir� executar at� primeira linha vazia � � da vari�vel linha, coluna 1.
    Loop
End Sub

'ATIVAR E DESATIVAR MOVIMENTO DE TELA
Sub at_des_Mov_tela()
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
End Sub

'ABRIR E FECHAR UM FORMUL�RIO
Sub abrir_fechar_userform()
    Nomedoformul�rio.Show
    Unload nome_do_formulario
End Sub

'COMANDOS PARA MOVIMENTAR COM SETA TECLADO
Sub Setas_Teclado()
    Selection.End(xlToLeft).Select '(esquerda)
    Selection.End(xlToRight).Select '(Direita)
    Selection.End(xlToup).Select '(Baixo)
    Selection.End(xlTodown).Select '(Cima)
End Sub

'DESLOCAR X C�LULAS PARA O LADO COLULA, PARA LADOS DAS LINHAS
Sub deslocar_celulas()
     ActiveCell.Offset(0, -4) = Range("A1")
     'OU
     ActiveCell.Offset(1, 0) = Worksheets("BASE DINAMICA").Range("A2").Value

End Sub

'ATIVAR E DESATIVAR ABAS DE UMA PLANILHA / FORMUL�RIOS E MENUS DO EXCEL
Sub Abas_Activate_des()
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayWorkbookTabs = False
    
    Application.DisplayFormulaBar = True
    Application.DisplayFormulaBar = False
    
    Application.DisplayFullScreen = True
    Application.DisplayFullScreen = False

End Sub

'CONDI��O IF (SE)
Sub If_Se()
If "Escreva aqui a Condi��o" = True Then
    'c�digo aqui
Else
    'c�digo aqui
End If
End Sub

'PREENCHER CAIXAS DE COMBINA��O COM TEXTOS (DEVER� CRIAR VALIDA��O DE DADOS NAS C�LULAS ONDE CONSTAM OS VALORES.)
Sub preencher_combobox()
    ultima_linha = Sheets("NOME DA PLANILHA").Range("A1").End(xlDown).Row
    caixa_atividade.RowSource = "NOMEPLANILHA!A2:B" & ultima_linha  '(COLOCAR COMBOBOX DE ACORDO COM PROJETO)
    caixa_atividade2.RowSource = "NOMEPLANILHA!A2:B" & ultima_linha '(COLOCAR COMBOBOX DE ACORDO COM PROJETO)
    caixa_atividade3.RowSource = "NOMEPLANILHA!A2:B" & ultima_linha '(COLOCAR COMBOBOX DE ACORDO COM PROJETO)
    
    'OU
    
    LINHA = Sheets("Controle_de_Produtos").Range("A1048576").End(xlUp).Row
    caixa_produto.RowSource = "Controle_de_Produtos!B2:B" & LINHA
End Sub

'ESCONDER PLANILHA � ABRIR SOMENTE FORMS
Sub abrir_somente_forms()
    Application.Visible = False
    CADASTRO.Show
End Sub

'ENCONTRAR VALOR DO N�MERO DA �LTIMA LINHA VAZIA
Sub ultima_linha_vazia()
    LINHA = Range("A1").End(xlDown).Row + 1
End Sub

'LIMPAR C�LULAS DE ACORDO COM VALORES ISERIDOS
Sub limpar_celulas()
    Range("A3:E250").Select
    Application.CutCopyMode = False
    Selection.ClearContents
End Sub

'ATIVAR E DESATIVAR ALERTAS DO EXCEL, TEMOS COMO EXEMPLO PERGUNTANDO SE DESEJA REALMENTE SAIR DE UM ARQUIVO SEM SALVAR
Sub alertas()
    Applicatiion.DisplayAlerts = False
    Applicatiion.DisplayAlerts = True
End Sub

' TECLA DE ENTER / ESPA�O
Sub enter_espa�o()
    SendKeys "{enter}", True
    SendKeys " "
End Sub

'COLOCAR DATA NO FORMATO CORRETO (BRASIL)
Sub data_format()
    Cells(LINHA, 3) = VBA.Format(TextBox1.Value, "mm/dd/yy")
End Sub

'---------------------------------------------------------------------------------
'MODELOS DE DECLARA��O DE VARI�VEL
'�  Texto
'Dim texto As String
'� N�mero
'Dim numero As Integer
'�   N�mero Decimal
'Dim numero_decimal As Double
'�   N�meros Longos
'Dim numero_longo As Long
'
'�   Exemplo de Declara��o
'nome = Cells(1, 2).Value
'numero = Cells(1, 2).Value
'numero_decimal = Cells(1, 2).Value
'numero_grande = Cells(1, 2).Value
'
'�   Declara��o de Abas
'Dim plan2 As Object
'Set aba_secundaria = sheets(�Segunda Aba�)
'�   Exemplo de Declara��o
'nome = aba_secundaria.Cells(1, 2).Value
'----------------------------------------------------------------------------------

'BUSCAR NOME DO COMPUTADOR
Sub nome_maquina()
    anexo = Application.UserName
    Planilha3.Cells(1, 1).Value = anexo
    '*Criar vari�vel
End Sub

'RETIRAR FILTRO DA PLANILHA
Sub retirar_filtro()
    Sheets("Planilha1").AutoFilterMode = False
End Sub

'INCLUIR FILTRO NA PLANILHA
Sub filtrar()
Sheets("Planilha1").UsedRange.AutoFilter 1, "Wasley"
'*N�mero 1 quer dizer o primeiro Filtro
'*Nome � o que deve ser filtrado
End Sub

'LIMPARA DADOS DA PLANILHA ( *LIMPAR� TODAS AS C�LULAS USADAS  DA PLANILHA ESPEC�FICA)
Sub limpar_tudo()
    Sheets("Planilha1").UsedRange.Clear
End Sub

'SELECIONA E COPIA TODO ESPA�O UTILIZADO NA PLANILHA (*INCLUSIVE C�LULAS EM BRANCO)
Sub copiar_tudo()
    Sheets("Planilha1").UsedRange.Copy
End Sub

'AJUSTAR TAMANHO DA COLUNA AUTOMATICAMENTE / CENTRALIZAR DADOS DE UMA C�LULA
Sub ajustar_coluna()
    Sheets("Planilha1").Columns.AutoFit
    Sheets("Planilha1").Rows.HorizontalAlignment = xlHAlignCenter
End Sub

'SALVAR PLANILHA
Sub Salvar_planilha()
    ThisWorkbook.Save
End Sub

'COLOCAR ID PLANILHA
Sub colocar_id()
    LINHA = Sheets("Planilha1").Range("A1000000").End(xlUp).Row + 1
    Sheets("Planilha1").Cells(LINHA, 1).Value = WorksheetFunction.Max(Sheets("Planilha1").Range("A:A")) + 1
End Sub

'LOCALIZAR OUTRO VALOR NA PLANILHA DE ACORDO COM UM VALOR DENTRO DE UMA TEXBOX
Sub localizar_por_texbox()
    LINHA = Sheets("Planilha1").Range("A:A").Find(TextBox1.Value).Row
End Sub

'SE UMA TECLA FOR PRESSIONADA. - ATEN��O CADA N�MERO EQUIVALEM A UMA TECLA, DEVEMOS CONSULTAR TABELAS.
Sub TeclaKey()
     If KeyCode = 48 Then
        MSGBOX "Fun��o Excel"
    End If
End Sub


