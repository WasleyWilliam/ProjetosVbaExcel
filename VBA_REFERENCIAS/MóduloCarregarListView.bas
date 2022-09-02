Attribute VB_Name = "M�duloCarregarListView"
'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       - Carregam Informa��o do Access para o ListView
'       - Conecta, desconecta do BD
'       - Fecha o RecordSet
'       -
'==================================================================================================================
'==================================================================================================================
                                        'AUTOR.........:WASLEY WILLIAM
                                        'CONTATO.......:CHAVE D1G3
                                        'DESCRI��O.....:CABECALHO LISTVIEW
                                        'REFERENCIA....:
'==================================================================================================================
Sub Carregar_Dados_Listview()
On Error GoTo Erro
Dim lista As Variant

Set rs = New ADODB.Recordset
M�dulo1.ConectarBD '(Chama o M�dulo de conectar o BD)

'                        *EM NOME DA TABELA DEVE TER O NOME IGUAL A TABELA DO BD
rs.Open "SELECT * FROM (NOME DA TABELA) ", Conexao, adOpenKeyset, adLockReadOnly

While Not rs.EOF
    With UserForm2.ListView1 '(ALTERAR - UserForm2.ListView1)
    Set lista = UserForm2.ListView1.ListItems.Add(Text:=rs(0)) '(ALTERAR - UserForm2.ListView1)
                lista.ListSubItems.Add Text:=rs(1)
                lista.ListSubItems.Add Text:=rs(2)
                lista.ListSubItems.Add Text:=rs(3)
                lista.ListSubItems.Add Text:=rs(4)
                lista.ListSubItems.Add Text:=rs(5)
                lista.ListSubItems.Add Text:=rs(6)
'                                         *(1,2,3,4,5,6)Refere-se a quantidade de colunas no BD
        End With
        rs.MoveNext
Wend
M�dulo1.Fechar_Rs '(Chama o M�dulo de fechar o BD)
M�dulo1.DesconectarBD '(Chama o M�dulo de desconectar o BD)
Exit Sub
Erro:
MsgBox "CARREGAR DADOS ! REFERENCIA....:M�DULO 0001", vbCritical, "SALVAR"
End Sub
'==================================================================================================================
                                        'FINAL DO C�DIGO
'==================================================================================================================
