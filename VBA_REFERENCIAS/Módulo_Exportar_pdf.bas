Attribute VB_Name = "Módulo_Exportar_pdf"
'==================================================================================================================
'   O(s) Código(s) abaixo executa(m) a(s) seguinte(s) função(ões):
'       -Exportar Dados em PDF
'       -
'       -
'       -
'==================================================================================================================
'==================================================================================================================
                                          'AUTOR.........:WASLEY WILLIAM
                                          'CONTATO.......:ww.adm@outlook.com
                                          'DESCRIÇÃO.....:EXPORTAR DADOS EM PDF
                                          'REFERENCIA....:
'==================================================================================================================
Sub exportar_pdf()
        caminho = "C:\Users\D1g3\OneDrive - PETROBRAS\GPI\Relatório" & Date & ".pdf"
        Sheets("Planilha1").ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=True

End Sub



'==================================================================================================================
                                        'FINAL DO CÓDIGO
'==================================================================================================================
