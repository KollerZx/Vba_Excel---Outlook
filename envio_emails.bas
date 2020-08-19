Sub enviar_email_SMA()

Dim OutApp As Object
Dim Email As Object

'Verifica se o Outlook está aberto e caso não esteja abre a aplicação
On Error Resume Next
Set Email = GetObject(, "Outlook.Application")

If Email Is Nothing Then
Set OutApp = CreateObject("Outlook.Application")
End If
'Declaração das variaveis que serão referência para as colunas da planilha
nome = 1
secretaria = 2
seção = 3
endereço = 4
ano = 5
periodo = 6
prazo = 7

'variavel que indica a linha de inicio dos valores a serem considerados da planilha 
linha = 6

'Pega a posição da ultima linha preenchida da coluna A e define como limite para nosso loop
linha_fim = Range("A2").End(xlDown).Row

'Loop para repetição da criação e envio dos e-mails com o anexo enquanto houver valores na coluna
While linha <= linha_fim

    'Cria um Email
    Set Email = OutApp.createItem(0)
    
        'define a variavel anexo como uma string que irá conter o caminho (Path) dos arquivos
        Dim anexo As String
        
       'Condição criada pois dentro de uma divisão havia mais de um arquivo para endereços diferentes
        If Cells(linha, nome).Value = "R.H" Then
            'a variavel anexo irá receber o caminho do arquivo a ser anexado
            anexo = ThisWorkbook.Path
            anexo = anexo & "\" & Cells(linha, ano) & "\" & Cells(linha, periodo) & "\" _
            & Cells(linha, secretaria) & "\" & Cells(linha, nome) & "\" & Cells(linha, seção) & ".pdf"
        
        'Caso o valor da celula não seja "R.H" o caminho padrão dos arquivos se mantém
        Else
            anexo = ThisWorkbook.Path
            anexo = anexo & "\" & Cells(linha, ano) & "\" & Cells(linha, periodo) & "\" _
            & Cells(linha, secretaria) & "\" & Cells(linha, nome) & "\" & Cells(linha, nome) & ".pdf"
        End If
        
        'Configuração do Email
        With Email
            'mostra janela de Email
            .display
            'Destinatário do E-mail
            .to = Cells(linha, endereço).Value
            'Para envio com Cópia para outro endereço
            '.CC = ("")
            'Para envio com cópia oculta
            '.BCC = ("")
            'Assunto do Email
            .Subject = "Espelho de Ponto " & Cells(linha, periodo)
            'Corpo do Email
            .Body = "Boa tarde," & Chr(10) & Chr(10) _
            & "Segue os espelhos da competência " & Cells(linha, periodo).Value & "." & Chr(10) _
            & "Por gentileza verificar e nos enviar até dia: " & Cells(linha, prazo).Value & "." & Chr(10) _
            & "Na falta de algum espelho, me enviar através deste e-mail a solicitação dos faltantes." & Chr(10) _
            & "Qualquer dúvida, estou à disposição" _
            & Chr(10) & Chr(10) & _
            "Att,"
            'Para anexar arquivos no E-mail
            .Attachments.Add anexo
            'Solicita confirmação da entrega
            .OriginatorDeliveryReportRequested = True
            'Solicita confirmação de leitura
            .ReadReceiptRequested = True
            'Envia o Email
            .Send
        End With
    'Incrementa a variável linha e segue o loop
    linha = linha + 1
Wend
End Sub




