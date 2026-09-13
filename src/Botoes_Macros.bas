Attribute VB_Name = "Botoes_Macros"
' ---------------------------------------------------------------------------
'  CRIA UM PAINEL DE BOTÕES AUTOMÁTICO NA PLANILHA (RODE UMA VEZ)
' ---------------------------------------------------------------------------
Public Sub Criar_Painel_De_Botoes()
    Dim ws As Worksheet
    Dim btn1 As Shape, btn2 As Shape, btn3 As Shape, btn4 As Shape
    
    Set ws = ThisWorkbook.Worksheets("SendBulkEmails")
    ws.Activate
    
    ' Insere linhas no topo se necessário
    If ws.Range("A1").Value <> "" Then
        ws.Rows("1:3").Insert Shift:=xlDown
    End If
    
    ' Botão 1: Revisar E-mails (Azul)
    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 10, 140, 35)
    With btn1
        .TextFrame.Characters.Text = "Revisar E-mails"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(0, 102, 204)
        .OnAction = "Revisar_Emails_Em_Andamento"
    End With
    
    ' Botão 2: Enviar E-mails (Verde)
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 170, 10, 140, 35)
    With btn2
        .TextFrame.Characters.Text = "Enviar E-mails"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(46, 139, 87)
        .OnAction = "Enviar_Emails_Em_Andamento"
    End With
    
    ' Botão 3: Configurar Horário (Cinza Escuro)
    Set btn3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 320, 10, 150, 35)
    With btn3
        .TextFrame.Characters.Text = "Configurar Horário"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(90, 90, 90)
        .OnAction = "Configurar_Horario_E_Dia"
    End With
    
    ' Botão 4: Limpar Log (Laranja)
    Set btn4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, 480, 10, 130, 35)
    With btn4
        .TextFrame.Characters.Text = "Limpar Log"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(218, 112, 214)
        .OnAction = "Limpar_Log_Para_Reenviar"
    End With
    
    MsgBox "Painel de botões criado com sucesso no topo da planilha!", vbInformation, "Botões Prontos"
End Sub