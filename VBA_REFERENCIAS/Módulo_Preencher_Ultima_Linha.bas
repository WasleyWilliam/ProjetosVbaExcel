Attribute VB_Name = "M�dulo_Preencher_Ultima_Linha"
Sub Preencher_ultima_Linha()

'=========================================================
'Os comandos abaixo possuem como fun��o:
    'Encontrar �ltima Linha (Faz um UP da C�lula 50000 para cima)+1
    'Preenche na �ltima linha os dados que constam no texbox1
'=========================================================



'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....: ENCONTRAR E PREENCHER A �LTIMA LINHA VAZIA
'REFERENCIA....:

Planilha5.Select
Linha = Range("A50000").End(xlUp).Row + 1
Cells(Linha, 1).Value = "* " & UserForm1.TextBox1.Value

'=========================================================
End Sub
    
    
    
    
    
    
    
    
    
