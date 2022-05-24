#If Win64 Then
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongLong, ByVal nCmdShow As LongLong) As Long
#Else
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Public IE As Object
Public Const SW_MAXIMIZE = 3
Function xWait()
   Do While IE.Busy Or IE.ReadyState <> 4
       DoEvents
   Loop
End Function

Sub updateProceduralClasses()

Dim User                As String
Dim Pass                As String
Dim Cod                 As String
Dim nLastRow            As Integer

On Error GoTo ErrorHandler

   nLastRow = P1_HOME.Range("D" & Rows.Count).End(xlUp).Row
   P1_HOME.[E1:E999].Font.Color = vbBlack
   P1_HOME.[E1:E999].Font.Italic = False
   P1_HOME.[E:E].HorizontalAlignment = xlLeft
   
   If P1_HOME.[D2] = "" Then
       MsgBox "Nenhum Código disponível na Aba 'Home', favor preencher os dados", vbExclamation, ""
       Exit Sub
   End If
   
   Set IE = CreateObject("InternetExplorer.Application")
   ShowWindow IE.hwnd, SW_MAXIMIZE
   IE.Navigate "https://esaj.tjce.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjce.jus.br%2Fesaj%2Fj_spring_cas_security_check"
   IE.Visible = True
   
   xWait
   
   With IE.Document.All
       If .Item("identificacao").innerText <> "CICERO EDIVAN OLIVEIRA LIMA  (Sair)" Then
           .Item("linkAbaCpf").Click
           .Item("usernameForm").Value = P1_HOME.Range("B2")
           .Item("passwordForm").Value = P1_HOME.Range("B3")
           IE.Document.getElementsByClassName("spwBotaoDefault ")(0).Click
           Application.Wait Now + TimeValue("00:00:02")
           xWait
       End If
   End With
   
   IE.Navigate "https://esaj.tjce.jus.br/cpopg/open.do?gateway=true"
   
   xWait
   
   For nX = 2 To nLastRow
   On Error GoTo ErrorHandler2
       If P1_HOME.Cells(nX, "D") <> "" And P1_HOME.Cells(nX, "E") = "" Then
           If fExistCodigo(P1_HOME.Cells(nX, "D")) Then
               With IE.Document.All
                   Cod = Mid(P1_HOME.Range("D" & nX).Text, 1, 15)
                   Cod2 = P1_HOME.Range("D" & nX)
                   .Item("cbPesquisa").SelectedIndex = 0
                   .Item("numeroDigitoAnoUnificado").innerText = Cod
                   .Item("foroNumeroUnificado").innerText = Right(P1_HOME.Range("D" & nX).Text, 4)
                   .Item("select2-chosen-1").innerText = "Todos os foros"
                   .Item("numeroDigitoAnoUnificado").Focus
                   Application.SendKeys "0"
                   xWait
                   .Item("botaoConsultarProcessos").Click
               End With
               xWait
               Clausula = IE.Document.getElementByID("classeProcesso").innerText
               dataHora_processo = IE.Document.getElementByID("dataHoraDistribuicaoProcesso").innerText
               If Clausula <> "" Then P1_HOME.Range("E" & nX) = Clausula: P1_HOME.Range("F" & nX) = dataHora_processo
                       
               IE.Navigate "https://esaj.tjce.jus.br/cpopg/open.do?gateway=true"
               xWait
                           
           Else
               P1_HOME.Range("E" & nX) = "Duplicado"
           End If
       End If
1    Next

   Set IE = Nothing

ErrorHandler:
   If Trim(nX) <> "" Then
        P1_HOME.Range("E" & nX) = " * processo em segredo de justica - necessita senha * "
        IE.Navigate "https://esaj.tjce.jus.br/cpopg/open.do?gateway=true"
        xWait
        Resume 1
   Else
        MsgBox "Erro detectado!" & vbNewLine & vbNewLine & "Contate o administrador do sistema.", vbCritical, "ERROR"
   End If
   
   Set IE = Nothing
   Application.ScreenUpdating = True
End Sub