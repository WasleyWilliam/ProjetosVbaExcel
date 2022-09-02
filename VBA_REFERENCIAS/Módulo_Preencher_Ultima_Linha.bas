Attribute VB_Name = "Módulo_Preencher_Ultima_Linha"
Sub Preencher_ultima_Linha()

'=========================================================
'Os comandos abaixo possuem como função:
    'Encontrar última Linha (Faz um UP da Célula 50000 para cima)+1
    'Preenche na última linha os dados que constam no texbox1
'=========================================================



'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRIÇÃO.....: ENCONTRAR E PREENCHER A ÚLTIMA LINHA VAZIA
'REFERENCIA....:

Planilha5.Select
Linha = Range("A50000").End(xlUp).Row + 1
Cells(Linha, 1).Value = "* " & UserForm1.TextBox1.Value

'=========================================================
End Sub
    
    
    
    
    
    
    
    
    
