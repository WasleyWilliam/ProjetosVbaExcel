Attribute VB_Name = "MóduloCriar_Um_txt"
Sub Criar_txt_com_dados_planilha()
'---------------------------------------
' Os Comandos abaixo executam as seguintes funções:
'     - Cria um arquivo .TXT
'     - Abre o arquivo criado
'          * No lugar do V... deve ser informado o endereço onde o arquivo será salvo
'
'---------------------------------------


'----------------------------------------
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:CHAVE D1G3
'DESCRIÇÃO.....:CRIA E ABRE UM ARQUIVO DE
'               BLOCO DE NOTAS COM TEXTOS DE UMA TABELA
'---------------------------------------

endereco = "V:...\NomeDoArquivo.txt"

Open endereco For Output As 1
Planilha5.Select
Range("A2").Select
    Do While ActiveCell.Value <> ""
        Print #1, ActiveCell.Value
        Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
    Loop
Close 1

'As linhas abaixo abrem o arquivo criado
Dim objShell As Object
Set objShell = CreateObject("Shell.Application")
objShell.Open (endereco)


End Sub
'--------------------------------------


