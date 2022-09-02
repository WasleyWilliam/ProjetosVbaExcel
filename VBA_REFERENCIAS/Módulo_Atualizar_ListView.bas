Attribute VB_Name = "Módulo_Atualizar_ListView"


'==================================================================================================================
'   O(s) Código(s) abaixo executa(m) a(s) seguinte(s) função(ões):
'       - Chama Módulo Conectar/Desconectar BD
'       - Abre/Fecha o RecordSet
'       - Atualiza a ListView
'       - Tratamento de Erro
'==================================================================================================================
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRIÇÃO.....:ATUALIZAR LISTVIEW NO BD ACCESS
                                          'REFERENCIA....:
'==================================================================================================================
Sub ATUALIZAR_LISTVIEW()
On Error GoTo Erro
Dim lista As Variant
UserForm2.ListView1.ListItems.Clear '(MODIFICAR NOME DO LISTVIEW)
Set rs = New ADODB.Recordset
Módulo1.ConectarBD
                    '*O NOME A SUBSTITUIR "NOME_DA_PLANILHA" DEVE SER EXATAMENTE O MESMO DA PLANILHA DO BD ACCESS
rs.Open "SELECT * FROM NOME_DA_PLANILHA ", Conexao, adOpenKeyset, adLockReadOnly

While Not rs.EOF
    With UserForm2.ListView1 '(MODIFICAR NOME DO LISTVIEW)
    Set lista = UserForm2.ListView1.ListItems.Add(Text:=rs(0)) '(MODIFICAR NOME DO LISTVIEW)
                lista.ListSubItems.Add Text:=rs(1)
                lista.ListSubItems.Add Text:=rs(2)
                lista.ListSubItems.Add Text:=rs(3)
                lista.ListSubItems.Add Text:=rs(4)
                lista.ListSubItems.Add Text:=rs(5)
                lista.ListSubItems.Add Text:=rs(6)
                                 'Texto Cabeçalho, Tamanho Cabeçalho, Alinhamento
        
        End With
        rs.MoveNext
Wend
Módulo1.Fechar_Rs
Módulo1.DesconectarBD
Exit Sub
Erro:
MsgBox "CARREGAR DADOS ! REFERENCIA....:MÓDULO  ", vbCritical, "SALVAR"
End Sub
'==================================================================================================================
                                                'FINAL DO CÓDIGO
'==================================================================================================================
