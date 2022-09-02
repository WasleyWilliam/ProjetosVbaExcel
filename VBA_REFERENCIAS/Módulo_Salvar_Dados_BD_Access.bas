Attribute VB_Name = "M�dulo_Salvar_Dados_BD_Access"


'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       - Busca M�dulo de Conectar/Desconecta BD
'       - Abre/Fecha RecordSet
'       - Busca Dados Repetidos
'       - Salva dados no BD
'       - Executa tratamento de erro
'==================================================================================================================
'==================================================================================================================
                                            'AUTOR.........:WASLEY WILLIAM
                                            'CONTATO.......:CHAVE D1G3
                                            'DESCRI��O.....:SALVAR DADOS NO BANCO DE DADOS ACCSSES
                                            'REFERENCIA....:
'==================================================================================================================
Sub Salvar_Dados_BD_Access()
On Error GoTo Erro
Set rs = New ADODB.Recordset
ConectarBD                      '*ATEN��O EM NOME EXATO DA TABELA DEVE-SE RETIRAR AS () E COLOCAR EXATAMENTE O NOME DA TABELA NO BD ACCESS
rs.Open "SELECT * FROM (NOME EXATO DA TABELA)", Conexao, adOpenKeyset, adLockPessimistic

'-------------------------------------------- BUSCAR REPETIDAS------------------------------------------------------
Do While Not rs.EOF
            If rs!NOME_DA_COLUNA = "" & UserForm2.TextBox1.Text Then '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS) , (MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
                    MsgBox "PADR�O J� CADASTRADO!", vbExclamation, "SALVAR" '(MENSAGEM DE DADOS REPETIDO)
                    If Not rs Is Nothing Then
                    rs.Close '(FECHANDO RECORDSET)
                    Set rs = Nothing
                    DesconectarBD '(FECHANDO O BANCO DE DADOS)
                    Exit Sub
            End If
            End If
rs.MoveNext
Loop
'---------------------------------------------------------------------------------------------------------------------
rs.AddNew '(NOVO RECORDSET)

'PREENCHIMENTO FORMUL�RIO
rs!REFERENCIA = UserForm2.TextBox2.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!PALAVRA_CHAVE = UserForm2.TextBox78.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!DESCRICAO = UserForm2.TextBox4.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!UNIDADE_OU_TAG = UserForm2.TextBox3.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)

'PREENCHIMENTO AUTOM�TICO
rs!DATA_HORA = VBA.Date & " - " & Time '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS),
anexo = Application.UserName 'CRIANDO VARI�VEL COM NOME DA M�QUINA
rs!INCLUIDO_POR = anexo '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS),(ANEXO � O NOME DA M�QUINA
rs.Update
MsgBox UserForm2.Label2.Caption & " ADICIONADO COM SUCESSO!", vbInformation, "SALVAR" '(MENSAGEM DE DADOS ADICIONADO COM SUCESSO)

UserForm2.TextBox1.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox2.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox2.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox4.Text = Empty '(LIMPA DADOS DA TEXBOX)

If Not rs Is Nothing Then
        rs.Close '(FECHANDO RS)
        Set rs = Nothing
End If
DesconectarBD '(DESCONECTANDO BD)
Exit Sub
Erro:
MsgBox "SALVAR! REFERENCIA....: ", vbCritical, "SALVAR"
End Sub
'==================================================================================================================
                                        'FINAL DO C�DIGO
'==================================================================================================================


