Attribute VB_Name = "Módulo_EnvioDe_Email"
Sub envio_de_email()

destinatario = ""
ComCopia = ""
assunto = ""
Set Email = objeto_outlook.createitem(0)
        Email.display
        Email.to = Cells(LINHA, 13).Value
        Email.cc = "sferraz@petrobras.com.br"
        Email.Subject = "APRESENTAÇÃO DE DDS"
        Email.body = Cells(LINHA, 4).Value & ", " & Cells(1, 20).Value & Chr(10) & Chr(10) _
        & "Sua Apresentação de DDS - "" & Cells(linha, 3).Value & Chr(10) & Chr(10)" _
        & "Está Agendada"" & Chr(10) & Chr(10)" _
        & Cells(LINHA, 5).Value & ", " & Cells(LINHA, 6).Value & Chr(10) & Chr(10) _
        & "Hora de Início " & VBA.Format(Cells(LINHA, 9).Value, "hh:mm") & " " & Chr(10) & Chr(10) _
        & "Atenciosamente" & Chr(10) & "Wasley William"
        Email.send
        Cells(LINHA, 12) = "E-mail Enviado em " & VBA.Format(Cells(1, 18).Value, "DD/MM/YYYY") & " ás " & VBA.Format(Cells(1, 18).Value, "hh:mm")

End Sub

