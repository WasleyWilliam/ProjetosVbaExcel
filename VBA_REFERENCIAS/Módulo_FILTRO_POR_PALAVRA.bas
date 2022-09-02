Attribute VB_Name = "M�dulo_FILTRO_POR_PALAVRA"
'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       - Filtra os dados em uma listview de acordo com uma textbox do projeto
'       - Chama o M�dulo Conectar / Desconecta BD
'       - Abre e Fecha o RecordSet
'       - O c�dico coloca na c�lula (1,1) o valor de "NENHUM CRIT�RIO ENCONTRADO" OU ENCONTRADO PALAVRA CHAVE" com isso � feito um if de compara��o ao chamar o c�digo
'==================================================================================================================
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRI��O.....: FILTRO DE DADOS DO BD ACCESS
                                          'REFERENCIA....:
'==================================================================================================================

Sub FILTRO_POR_PALAVRA()
UserForm2.ListView1.ListItems.Clear '(MODIFICAR O NOME PARA A LISTVIEW DO PROJETO - PROCESSO LIMPAR� LISTVIEW)
Dim SQL As String
Dim lista As Object

Set rs = New ADODB.Recordset

M�dulo1.ConectarBD
'                        (MODIFICAR O NOME DA PLANILHA A PROCURAR - ATEN��O, TEM QUE MANTER UM ESPA�O ANTES DAS ")
SQL = "SELECT * FROM NOME_DA_PLANILHA "
SQL = SQL & "WHERE NOME_DA_COLUNA Like '%" & UserForm2.TextBox1.Text & "%'"
'                        '(MODIFICAR O NOME DA COLUNA A PROCURAR E TAMB�M INDICAR O SEU TEXBOX COM CRIT�RIO DE PESQUISA ")

rs.Open SQL, Conexao, adOpenKeyset, adLockReadOnly

If rs.RecordCount = 0 Then

Planilha1.Cells(1, 1).Value = "NENHUM CRIT�RIO ENCONTRADO"
   
    M�dulo1.Fechar_Rs
    M�dulo1.DesconectarBD
    Exit Sub
End If
While Not rs.EOF

With UserForm2.ListView1 '(MODIFICAR O NOME PARA A LISTVIEW DO PROJETO)
    Set lista = UserForm2.ListView1.ListItems.Add(Text:=rs(0))
                lista.ListSubItems.Add Text:=rs(1)
                lista.ListSubItems.Add Text:=rs(2)
                lista.ListSubItems.Add Text:=rs(3)
                lista.ListSubItems.Add Text:=rs(4)
                lista.ListSubItems.Add Text:=rs(5)
                lista.ListSubItems.Add Text:=rs(6)

    End With
    rs.MoveNext
Wend

M�dulo1.Fechar_Rs
M�dulo1.DesconectarBD
Set lista = Nothing
Planilha1.Cells(1, 1).Value = "ENCONTRADO PALAVRA CHAVE"
End Sub
'==================================================================================================================
                                                'FINAL DO C�DIGO
'==================================================================================================================
