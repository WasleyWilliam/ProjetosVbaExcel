Attribute VB_Name = "M�dulo_Preencher_Txb_ListView"
'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       - Carrega os dados do BD Access para os texboxs do projeto
'       - Tratamento de Erro
'       -
'       -
'==================================================================================================================
'==================================================================================================================
                                       'AUTOR.........:WASLEY WILLIAM
                                       'CONTATO.......:ww.adm@outlook.com
                                       'DESCRI��O.....:CARREGAR DADOS DO BD PARA UMA LISTVIEW
                                       'REFERENCIA....:
'==================================================================================================================
Sub Carregar_Listview_Textbox()

On Error GoTo Erro
Dim Linha As Double

With UserForm2.ListView1 '(MODIFICAR PELO NOME DA LISTVIEW DO PROJETO)
                Linha = .SelectedItem.Index
                UserForm2.TextBox5.Value = .ListItems(Linha).Text 'ID
                UserForm2.TextBox2.Value = .ListItems(Linha).ListSubItems(1).Text '(CADA TEXBOX DO PROJETO RECEBER� UM VALOR DE ITENS DA LISTA)
                UserForm2.TextBox3.Value = .ListItems(Linha).ListSubItems(2).Text
                UserForm2.TextBox4.Value = .ListItems(Linha).ListSubItems(4).Text
                UserForm2.TextBox78.Value = .ListItems(Linha).ListSubItems(3).Text
                
End With
Exit Sub
Erro:
MsgBox "NADA PARA CARREGAR, REFA�A A BUSCA!", vbInformation, "FILTRO"
End Sub
'==================================================================================================================
                                                'FINAL DO C�DIGO
'==================================================================================================================

