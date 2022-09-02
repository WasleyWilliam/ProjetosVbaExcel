Attribute VB_Name = "Módulo_Criar_Pastas_e_SubPastas"
'=========================================================
'O código a seguir executam as seguintes funções:
    '-Declara Variáveis
    '-Cria Pasta Principal
    '-Cria SubPastas baseado na pasta principal
        ' *Obs: em NPASTA = "C:..." o C... deve indicar o caminho de geração da pasta
        ' *UserForm2.TextBox2.Value equivale ao nome da pasta de um valor digitado em um texbox
'=========================================================
Sub CRIAR_PASTAS_SUBPASTAS()
'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRIÇÃO.....:CRIAR PASTAS E SUB PASTAS
'REFERENCIA....:
Dim NPASTA As String
Dim SUBPASTA As String

EstaPastaDeTrabalho.Activate
NPASTA = "C:..." & UserForm2.Label2.Caption

SUBPASTA = NPASTA & "\" & UserForm2.TextBox2.Value

        If Dir(NPASTA, vbDirectory) = "" Then
            MkDir NPASTA
        End If
        If Dir(SUBPASTA, vbDirectory) = "" Then
            MkDir SUBPASTA
        End If

End Sub
'=========================================================
