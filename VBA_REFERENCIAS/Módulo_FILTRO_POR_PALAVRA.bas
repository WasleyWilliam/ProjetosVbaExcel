Attribute VB_Name = "Módulo_FILTRO_POR_PALAVRA"
'==================================================================================================================
'   O(s) Código(s) abaixo executa(m) a(s) seguinte(s) função(ões):
'       - Filtra os dados em uma listview de acordo com uma textbox do projeto
'       - Chama o Módulo Conectar / Desconecta BD
'       - Abre e Fecha o RecordSet
'       - O códico coloca na célula (1,1) o valor de "NENHUM CRITÉRIO ENCONTRADO" OU ENCONTRADO PALAVRA CHAVE" com isso é feito um if de comparação ao chamar o código
'==================================================================================================================
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRIÇÃO.....: FILTRO DE DADOS DO BD ACCESS
                                          'REFERENCIA....:
'==================================================================================================================

Sub FILTRO_POR_PALAVRA()
UserForm2.ListView1.ListItems.Clear '(MODIFICAR O NOME PARA A LISTVIEW DO PROJETO - PROCESSO LIMPARÁ LISTVIEW)
Dim SQL As String
Dim lista As Object

Set rs = New ADODB.Recordset

Módulo1.ConectarBD
'                        (MODIFICAR O NOME DA PLANILHA A PROCURAR - ATENÇÃO, TEM QUE MANTER UM ESPAÇO ANTES DAS ")
SQL = "SELECT * FROM NOME_DA_PLANILHA "
SQL = SQL & "WHERE NOME_DA_COLUNA Like '%" & UserForm2.TextBox1.Text & "%'"
'                        '(MODIFICAR O NOME DA COLUNA A PROCURAR E TAMBÉM INDICAR O SEU TEXBOX COM CRITÉRIO DE PESQUISA ")

rs.Open SQL, Conexao, adOpenKeyset, adLockReadOnly

If rs.RecordCount = 0 Then

Planilha1.Cells(1, 1).Value = "NENHUM CRITÉRIO ENCONTRADO"
   
    Módulo1.Fechar_Rs
    Módulo1.DesconectarBD
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

Módulo1.Fechar_Rs
Módulo1.DesconectarBD
Set lista = Nothing
Planilha1.Cells(1, 1).Value = "ENCONTRADO PALAVRA CHAVE"
End Sub
'==================================================================================================================
                                                'FINAL DO CÓDIGO
'==================================================================================================================
