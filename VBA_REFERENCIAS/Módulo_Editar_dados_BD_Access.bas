Attribute VB_Name = "Módulo_Editar_dados_BD_Access"

'==================================================================================================================
'   O(s) Código(s) abaixo executa(m) a(s) seguinte(s) função(ões):
'       - Abre/Fecha RecordSet
'       - Chama Módulos de Conectar/Desconectar BD
'       - Salva dados no BD
'       - Executa tratamento de erro
'==================================================================================================================
'==================================================================================================================
                                            'AUTOR.........:WASLEY WILLIAM
                                            'CONTATO.......:ww.adm@outlook.com
                                            'DESCRIÇÃO.....:EDITAR DADOS DE UM BD ACCESS
                                            'REFERENCIA....:
'==================================================================================================================

Sub Editar_dados_BD_Access()
On Error GoTo Erro
Set rs = New ADODB.Recordset
ConectarBD
                            '*ATENÇÃO EM NOME EXATO DA TABELA DEVE-SE RETIRAR AS () E COLOCAR EXATAMENTE O NOME DA TABELA NO BD ACCESS
rs.Open "SELECT * FROM NOME DA PLANILHA WHERE ID=" & UserForm2.TextBox5.Text, Conexao, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then

    'rs!ID = TextBox4.TextBox1.Text
rs!REFERENCIA = UserForm2.TextBox2.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!PALAVRA_CHAVE = UserForm2.TextBox78.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!DESCRICAO = UserForm2.TextBox4.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
rs!UNIDADE_OU_TAG = UserForm2.TextBox3.Text '(MODIFICAR NOME DA COLUNA PELO NOME EXATO DA COLUNA NO BD ACCESS), '(MODIFICAR NOME DA TEXBOX PELO NOME DA SUA TEXBOX)
   
MsgBox "EDITADO COM SUCESSO!", vbInformation, "EDITAR"
     
UserForm2.TextBox1.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox2.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox2.Text = Empty '(LIMPA DADOS DA TEXBOX)
UserForm2.TextBox4.Text = Empty '(LIMPA DADOS DA TEXBOX)


UserForm2.TextBox5.Enabled = False

 rs.Update
Else

MsgBox "Não Encontrado!", vbCritical, "EDITAR"
End If

If Not rs Is Nothing Then

rs.Close
Set rs = Nothing
End If
DesconectarBD

Exit Sub
Erro:
MsgBox "SELECIONE AO LADO ALGUM PADRÃO! ", vbCritical, "REFERENCIA....:MÓDULO SIMPEP 0005"


End Sub

'==================================================================================================================
                                        'FINAL DO CÓDIGO
'==================================================================================================================
