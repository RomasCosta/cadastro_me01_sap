# cadastro_me01_sap
Cadastro de LOF-ME01 SAP


Sub CriarME01_Click()

    Dim resposta As VbMsgBoxResult
    
  ' Exibe uma caixa de mensagem com duas opções
    resposta = MsgBox("Escolha uma opção:" & vbCrLf & _
                      "Clique em Sim para LOF Única" & vbCrLf & _
                      "Clique em Não para LOF Dupla", vbYesNo + vbQuestion, "Escolha de Execução")
    
    ' Verifica a escolha do usuário
    If resposta = vbYes Then
        ' Aqui você pode colocar o código que deseja executar para a opção 1
        Call CriarME01_1
    Else
        ' Aqui você pode colocar o código que deseja executar para a opção 2
        Call CriarME01_2
    End If
End Sub

Sub CriarME01_1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOF")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim tableView As Object
    Dim j As Long
    Dim resposta As VbMsgBoxResult
    
    If Not IsObject(Appl) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
        
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
    session.FindById("wnd[0]").sendVKey 0
    
    
    For i = 2 To ultimaLinha
        
            On Error Resume Next
            
            'session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
    
            session.FindById("wnd[0]/usr/ctxtEORD-MATNR").Text = ws.Cells(i, 1).Value '"66-03416"
            session.FindById("wnd[0]/usr/ctxtEORD-WERKS").Text = "BR35"
            session.FindById("wnd[0]").sendVKey 0
            
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").Text = ws.Cells(i, 2).Value '"14.01.2025"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").Text = "31.12.9999"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").Text = ws.Cells(i, 3).Value '"103432"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").Text = "BR35"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-AUTET[10,0]").Text = "1"
   
            session.FindById("wnd[0]").sendVKey 11
    
            Application.Wait Now + TimeValue("00:00:02")
            
            statusMessage = session.FindById("wnd[0]/sbar").Text
            ws.Cells(i, "E").Value = statusMessage
            
            typeMessage = session.FindById("wnd[0]/sbar").messagetype
            ws.Cells(i, "F").Value = typeMessage
            
            If typeMessage <> "S" Then
                
                ws.Cells(i, "G").Value = "Erro"
                
                session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
                session.FindById("wnd[0]").sendVKey 0
                
            Else
                ws.Cells(i, "G").Value = "Sucesso"
                
            End If
         
    Next i
             
End Sub

Sub CriarME01_2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LOF")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim tableView As Object
    Dim j As Long
    
    If Not IsObject(Appl) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
        
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
    session.FindById("wnd[0]").sendVKey 0
    
    '--------------------------------------------------------------
    
    j = 3 '
    
    '--------------------------------------------------------------
    
    For i = 2 To ultimaLinha
        
            On Error Resume Next
            
            'session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
    
            session.FindById("wnd[0]/usr/ctxtEORD-MATNR").Text = ws.Cells(i, 1).Value '"66-03416"
            session.FindById("wnd[0]/usr/ctxtEORD-WERKS").Text = "BR35"
            session.FindById("wnd[0]/usr/ctxtEORD-WERKS").SetFocus
            session.FindById("wnd[0]/usr/ctxtEORD-WERKS").caretPosition = 4
            session.FindById("wnd[0]").sendVKey 0
            
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").Text = ws.Cells(i, 2).Value '"14.01.2025"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").Text = "31.12.9999"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").Text = ws.Cells(i, 3).Value '"103432"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").Text = "BR35"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-AUTET[10,0]").Text = "1"
            
            'ATIVAR PARA 2 FORNECEDORES
            
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,1]").Text = ws.Cells(j, 2).Value '"14.01.2025"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,1]").Text = "31.12.9999"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,1]").Text = ws.Cells(j, 3).Value '"103432"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,1]").Text = "BR35"
            session.FindById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-AUTET[10,0]").Text = "1"
   
            session.FindById("wnd[0]").sendVKey 11
    
            Application.Wait Now + TimeValue("00:00:02")
            
            statusMessage = session.FindById("wnd[0]/sbar").Text
            ws.Cells(i, "E").Value = statusMessage
            ws.Cells(j, "E").Value = statusMessage 'ATIVAR PARA 2 FORNECEDORES
            
            typeMessage = session.FindById("wnd[0]/sbar").messagetype
            ws.Cells(i, "F").Value = typeMessage
            ws.Cells(j, "F").Value = typeMessage 'ATIVAR PARA 2 FORNECEDORES
            
            If typeMessage <> "S" Then
                
                ws.Cells(i, "G").Value = "Erro"
                ws.Cells(j, "G").Value = "Erro" 'ATIVAR PARA 2 FORNECEDORES
                
                session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n me01"
                session.FindById("wnd[0]").sendVKey 0
                
            Else
                ws.Cells(i, "G").Value = "Sucesso"
                ws.Cells(j, "G").Value = "Sucesso" 'ATIVAR PARA 2 FORNECEDORES
                
            End If
            
        j = j + 2 'ATIVAR PARA 2 FORNECEDORES
            
        i = i + 1 'ATIVAR PARA 2 FORNECEDORES
         
    Next
             
End Sub





