Attribute VB_Name = "M�duloConectarADODB"

'=========================================================
'Os comandos abaixo possuem como fun��o:
    ' - Declara��o de Veri�veis
    ' - Conex�o com banco de dados ACCESS via ADODB
    ' - Desconex�o com banco de dados ACCESS
    ' - Fechar ADODB RecordSet
'=========================================================



'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....: VARI�VEIS DE CONEX�O COM BANCO DE DADOS ACCESS
'REFERENCIA....:
Public Conexao As ADODB.Connection
Public rs As ADODB.Recordset
'=========================================================
'=========================================================


Sub ConectarBD()
'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....:CONECTAR BANCO DE DADOS
'REFERENCIA....:
'=========================================================
On Error GoTo Erro
Dim StrConexao As String
Set Conexao = New ADODB.Connection
        'A Linha Abaixo deve ser indicado o "caminho" do arquivo access (C:...)
        StrConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\;Persist Security Info=False"
        Conexao.Open StrConexao

Exit Sub
Erro:
MsgBox "ERRO REFERENCIA....:M�DULO1 0002"
End Sub
'=========================================================
'=========================================================


Sub DesconectarBD()
'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....:DESCONECTAR BANCO DE DADOS
'REFERENCIA....:
'=========================================================

On Error GoTo Erro

    If Not Conexao Is Nothing Then
        Conexao.Close
        Set Conexao = Nothing
    End If
Exit Sub
Erro:
MsgBox "ERRO REFERENCIA....:M�DULO1 0003"
End Sub
'=========================================================
'=========================================================



Sub Fechar_Rs()
'=========================================================
'AUTOR.........:WASLEY WILLIAM
'CONTATO.......:ww.adm@outlook.com
'DESCRI��O.....:FECHAR RECORDSET
'REFERENCIA....:
'=========================================================
If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
    End If
End Sub
'=========================================================
'=========================================================

