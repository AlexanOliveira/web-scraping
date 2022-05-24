Sub updateProceduralClasses()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
On Error GoTo errHand

    P1_HOME.[E2:E999].Font.Color = vbBlack
    P1_HOME.[E2:E999].Font.Italic = False
    P1_HOME.[E:F].HorizontalAlignment = xlLeft
    P1_HOME.[E2:F999].ClearContents
    
    If Application.CountA(P1_HOME.[D2:D100]) = 0 Then
        MsgBox "Nenhum Código disponível na Aba 'Home', favor preencher os dados", vbExclamation, ""
        Exit Sub
    End If
    
    userName = CLI_USERNAME
    passWord = CLI_PASSWD
    
    Url = "https://esaj.tjce.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjce.jus.br%2Fesaj%2Fj_spring_cas_security_check"
    UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36"
    ContentType = "text/html; charset=utf-8"
    PostContentType = "application/x-www-form-urlencoded"
    
    Set doc = CreateObject("HTMLFile")
    
    With CreateObject("MSXML2.ServerXMLHTTP")
        .Open "GET", Url, False
        .send "username=" & CLI_USERNAME & "&password=" & CLI_PASSWD & "&lt=&_eventId=submit&pbEntrar=Entrar&signature=&certificadoSelecionado=&certificado="
        doc.body.innerHTML = .responseText
    
        DocumentCookie = .getResponseHeader("Set-Cookie")
        Execution = doc.querySelector("#formCertificado").elements.Execution.Value
    End With
    
    FormData = "username=" & userName & "&password=" & passWord
    FormData = FormData & "&lt=&execution=" & Execution & "&_eventId=submit&signature=&certificadoSelecionado=&certificado="
    
    With CreateObject("MSXML2.ServerXMLHTTP")
        .Open "POST", Url, False
        .setRequestHeader "Content-Type", PostContentType
        .setRequestHeader "User-Agent", UserAgent
        .send FormData
        doc.body.innerHTML = .responseText
    
        DocumentCookie = .getResponseHeader("Set-Cookie")
    End With
    
    For nX = 2 To P1_HOME.Range("D" & Rows.Count).End(xlUp).Row
        If P1_HOME.Cells(nX, "D") <> "" And P1_HOME.Cells(nX, "E") = "" Then
            If fExistCodigo(P1_HOME.Cells(nX, "D")) Then
                ProcessNumber = P1_HOME.Range("D" & nX)
    
                FormData = "conversationId=&cbPesquisa=NUMPROC"
                FormData = FormData & "&numeroDigitoAnoUnificado=" & Mid(ProcesNumber, 12, 4)
                FormData = FormData & "&forNumeroUnificado=" & Right(ProcessNumber, 4)
                FormData = FormData & "&dadoConsulta.valorConsulta=&dadosConsulta.tipoNuProcesso=UNIFICADO"

                With CreateObject("MSXML2.ServerXMLHTTP")
                    .Open "GET", "https://esaj.tjce.jus.br/cpopg/search.do?" & FormData, False
                    .setRequestHeader "Content-Type", PostContentType
                    .setRequestHeader "User-Agent", UserAgent
                    .setRequestHeader "Set-Cookie", DocumentCookie
                    .send
                    
                    doc.body.innerHTML = .responseText
                End With
            
                If IsObject(doc.getElementByID("classeProcesso")) Then
                    P1_HOME.Range("E" & nX) = doc.getElementByID("classeProcesso").innerText
                    P1_HOME.Range("F" & nX) = doc.getElementByID("dataHoraDistribuicaoProcesso").innerText
                ElseIf IsObject(doc.querySelector(".dropdown-menu.tooltip-campos")) Then
                    P1_HOME.Range("E" & nX) = " * processo em segredo de justica - necessita senha * "
                    P1_HOME.Range("E" & nX).Font.Color = RGB(225, 55, 65)
                    P1_HOME.Range("F" & nX) = Now
                ElseIf IsObject(doc.querySelector("mensagemRetorno")) Then
                    P1_HOME.Range("E" & nX) = doc.querySelector("mensagemRetorno").innerText
                    P1_HOME.Range("F" & nX) = Now
                End If
            Else
                P1_HOME.Range("E" & nX) = "Duplicado"
            End If
        End If
    Next

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Exit Sub

errHand:
   Application.Calculation = xlCalculationAutomatic
   Application.ScreenUpdating = True
   MsgBox "Um erro inexperado ocorreu. Entre em contato com AutenLab Suport!", vbCritical, "Erro:"
End Sub
