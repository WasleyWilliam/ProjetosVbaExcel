Attribute VB_Name = "M�dulo_Cabecalho_listview"

'==================================================================================================================
'   O C�digo abaixo executa a seguinte fun��o:
'       - Cria Cabe�alhos em uma ListView
'
'
'
'==================================================================================================================


'==================================================================================================================
'                                    AUTOR.........:WASLEY WILLIAM
'                                    CONTATO.......:CHAVE D1G3
'                                    DESCRI��O.....:CABE�ALHO LISTVIEW
'                                    REFERENCIA....:
'==================================================================================================================
Sub Cabecalho_ListView()

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
                                        'Texto Cabe�alho, Tamanho Cabe�alho, Alinhamento
        End With
End Sub
'==================================================================================================================
                                        'FINAL DO C�DIGO
'==================================================================================================================
