Sub enviar_email_marcacao_ponto_cc()
    Dim objeto_outlook As Object
    Set objeto_outlook = CreateObject("Outlook.Application")

    Dim Email As Object
    Dim ultima_linha As Long
    Dim i As Long

    ultima_linha = Cells(Rows.Count, 2).End(xlUp).Row 'Encontra a última linha da coluna B (nome)

    For i = 2 To ultima_linha 'Assume cabeçalho na linha 1
        If Trim(Cells(i, 17).Value) <> "" Then 'Coluna Q (17): só envia se houver observação
            Set Email = objeto_outlook.CreateItem(0)
            Email.To = Cells(i, 19).Value 'Coluna S (19): e-mail
            Email.CC = "fulano@empresa.com.br" 'Inclua aqui o(s) email(s) para CC
            Email.Subject = "Marcação de Ponto - Seniorx"
            Email.Body = "Olá, " & Cells(i, 2).Value & ";" & vbCrLf & vbCrLf & _
                "Estamos em fechamento do período e estão faltando marcações no seu ponto no sistema Seniorx." & vbCrLf & _
                Cells(i, 17).Value & vbCrLf & _
                "Atualize suas marcações!" & vbCrLf & vbCrLf & _
                "Muito obrigada." & vbCrLf & _
                "Atenciosamente"
            Email.Display 'Abre o e-mail para revisão manual no Outlook
        End If
    Next i
End Sub
