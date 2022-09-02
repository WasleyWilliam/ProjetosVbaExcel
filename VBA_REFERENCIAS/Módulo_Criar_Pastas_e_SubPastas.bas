Attribute VB_Name = "M�dulo_Criar_Pastas_e_SubPastas"
'=========================================================
'O c�digo a seguir executam as seguintes fun��es:
    '-Declara Vari�veis
    '-Cria Pasta Principal
    '-Cria SubPastas baseado na pasta principal
        ' *Obs: em NPASTA = "C:..." o C... deve indicar o caminho de gera��o da pasta
        ' *UserForm2.TextBox2.Value equivale ao nome da pasta de um valor digitado em um texbox
'=========================================================
'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....:CRIAR PASTAS E SUB PASTAS
'REFERENCIA....:


Sub CRIAR_PASTAS_SUBPASTAS()

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
