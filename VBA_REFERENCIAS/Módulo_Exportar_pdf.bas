Attribute VB_Name = "M�dulo_Exportar_pdf"
'==================================================================================================================
'   O(s) C�digo(s) abaixo executa(m) a(s) seguinte(s) fun��o(�es):
'       -Exportar Dados em PDF
'       -
'       -
'       -
'==================================================================================================================
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRI��O.....:EXPORTAR DADOS EM PDF
                                          'REFERENCIA....:
'==================================================================================================================
Sub exportar_pdf()
        caminho = "C:\Users\D1g3\OneDrive - PETROBRAS\GPI\Relat�rio" & Date & ".pdf"
        Sheets("Planilha1").ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=True

End Sub



'==================================================================================================================
                                        'FINAL DO C�DIGO
'==================================================================================================================
