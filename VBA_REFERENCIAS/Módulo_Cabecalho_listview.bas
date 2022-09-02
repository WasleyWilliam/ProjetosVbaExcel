Attribute VB_Name = "Módulo_Cabecalho_listview"
Sub Cabecalho_ListView()
'==================================================================================================================
'   O Código abaixo executa a seguinte função:
'       - Cria Cabeçalhos em uma ListView
'
'
'
'==================================================================================================================


'==================================================================================================================
'                                    AUTOR.........:WASLEY WILLIAM
'                                    CONTATO.......:CHAVE D1G3
'                                    DESCRIÇÃO.....:CABEÇALHO LISTVIEW
'                                    REFERENCIA....:
'==================================================================================================================

With UserForm2.ListView1 '(Alterar - UserForm2.ListView1)
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .MultiSelect = True
        .ColumnHeaders.Add Text:="ID", Width:=30, Alignment:=0
        .ColumnHeaders.Add Text:="REFERENCIA", Width:=65, Alignment:=0 '(Alterar - "REFERENCIA")
        .ColumnHeaders.Add Text:="PALAVRA_CHAVE", Width:=130, Alignment:=0 '(Alterar - "PALAVRA_CHAVE")
        .ColumnHeaders.Add Text:="DESCRICAO", Width:=450, Alignment:=0 '(Alterar - "DESCRICAO")
        .ColumnHeaders.Add Text:="DATA_HORA", Width:=100, Alignment:=0 '(Alterar - "DATA_HORA")
        .ColumnHeaders.Add Text:="INCLUIDO_POR", Width:=120, Alignment:=0 '(Alterar - "INCLUIDO_POR")
                                 'Texto Cabeçalho, Tamanho Cabeçalho, Alinhamento
End With
End Sub
'==================================================================================================================
                                        'FINAL DO CÓDIGO
'==================================================================================================================
