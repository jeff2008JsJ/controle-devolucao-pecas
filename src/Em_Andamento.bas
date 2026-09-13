Attribute VB_Name = "Em_Andamento"
Option Explicit

' ============================================================================
'  AVISO / SEGUIMENTO RPTG — EXCLUSIVO PARA STAS "EM ANDAMENTO" (COLETA)
'  - Regiões Sudeste divididas: Sudeste SP, Sudeste MG e Sudeste RJ/ES
'  - Leitura dinâmica dos e-mails pela aba Parametros
'  - Valores mantidos (R$) por mês e valor total informativo (sem débito)
'  - Filtro rigoroso: ignora estritamente peças "EM TRANSPORTE"
'  - Agendamento interno configurável (Dia e Horário)
'  - Assinatura: Nome do Analista (Qualidade - RPTG)
' ============================================================================

Private Const ABA_EMAILS As String = "SendBulkEmails"
Private Const ABA_BASE As String = "Base_RPTG"
Private Const ABA_PARAM As String = "Parametros"
Private Const STATUS_ALVO As String = "EM ANDAMENTO"

' Colunas na aba SendBulkEmails
Private Const COL_LOG As Long = 9                            ' Coluna I: Log do envio
Private Const COL_DETALHE As Long = 11                       ' Coluna K: Meses em aberto
Private Const COL_REGRA As Long = 12                         ' Coluna L: Status
Private Const COL_DEBITO As Long = 13                        ' Coluna M: Valor Total (R$)
Private Const COL_REGIAO As Long = 14                        ' Coluna N: Nome da Região
Private Const COL_CC As Long = 15                            ' Coluna O: Consultores em CC

' Células de configuração na aba Parametros
Private Const CEL_BANNER As String = "P6"
Private Const CEL_ATIVO As String = "B2"                     ' Envio Automático Ativo (SIM/NAO)
Private Const CEL_DIA As String = "B3"                       ' Dia da semana (2=Seg..6=Sex, 8=Diário)
Private Const CEL_HORA As String = "B4"                      ' Horário do disparo (ex: 09:30)

Private Const PARA_FIXO As String = "financeiro@exemplo.com.br"
Private Const COPIA_FIXA As String = "retorno.pecas@exemplo.com.br"
Private Const PULAR_JA_ENVIADOS As Boolean = True
Private Const PAUSA_SEGUNDOS As Double = 1                   ' Intervalo entre envios
Private Const MAX_REVISAR As Long = 5                        ' Quantos abrir na revisão (0 = todos)
Private Const DIAS_CICLO As Long = 7                         ' Ciclo semanal

Private Const ARQ_BANNER As String = "Assinatura_Empresa.png"
Private Const CAMINHO_BANNER As String = ""
Private Const LARGURA_BANNER As Long = 500
Private Const ARQ_LOG As String = "Log_Envio_Semanal_RPTG.txt"

Private silencioso As Boolean

' ---------------------------------------------------------------------------
'  1. AGENDAMENTO INTERNO AUTOMÁTICO (CHAMADO AO ABRIR O EXCEL)
' ---------------------------------------------------------------------------
Public Sub Verificar_E_Agendar_Envio()
    Dim diaSalvo As Integer, horaSalva As String
    Dim horaAtual As Double, horaAlvo As Double
    Dim hojeDia As Integer, ativo As String
    
    CriarAbaParametrosSeNaoExistir
    
    ativo = UCase(Trim(CStr(ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_ATIVO).Value)))
    If ativo = "NAO" Or ativo = "NÃO" Then Exit Sub
    
    diaSalvo = ObterDiaSalvo()
    horaSalva = ObterHoraSalva()
    
    If diaSalvo = 0 Or horaSalva = "" Then Exit Sub
    
    hojeDia = Weekday(Date)
    
    If diaSalvo = 8 Then
        If hojeDia < 2 Or hojeDia > 6 Then Exit Sub
    Else
        If hojeDia <> diaSalvo Then Exit Sub
    End If
    
    On Error Resume Next
    horaAlvo = TimeValue(horaSalva)
    horaAtual = TimeValue(Format(Now, "hh:nn:ss"))
    On Error GoTo 0
    
    If horaAtual >= horaAlvo Then
        Application.OnTime Now + TimeSerial(0, 0, 5), "Envio_Automatico_Em_Andamento"
    Else
        Application.OnTime Date + horaAlvo, "Envio_Automatico_Em_Andamento"
    End If
End Sub

' ---------------------------------------------------------------------------
'  2. CONFIGURAÇÃO DE HORÁRIO E TABELA DE REGIÕES (ALT + F8)
' ---------------------------------------------------------------------------
Public Sub Configurar_Horario_E_Dia()
    Dim wsP As Worksheet
    Dim diaEscolhido As String, horaEscolhida As String
    Dim numDia As Integer, horaValida As Date
    Dim textoDias As String
    
    CriarAbaParametrosSeNaoExistir
    Set wsP = ThisWorkbook.Worksheets(ABA_PARAM)
    
    textoDias = "Escolha o número do dia da semana para o envio automático:" & vbCrLf & vbCrLf & _
                "2 = Segunda-feira" & vbCrLf & _
                "3 = Terça-feira" & vbCrLf & _
                "4 = Quarta-feira" & vbCrLf & _
                "5 = Quinta-feira" & vbCrLf & _
                "6 = Sexta-feira" & vbCrLf & _
                "8 = Todos os dias úteis (Seg a Sex)"
                
    diaEscolhido = InputBox(textoDias, "Configurar Dia de Envio", ObterDiaSalvo())
    If Trim(diaEscolhido) = "" Then Exit Sub
    
    numDia = Val(diaEscolhido)
    If (numDia < 2 Or numDia > 6) And numDia <> 8 Then
        MsgBox "Opção inválida! Escolha um número de 2 a 6 ou 8.", vbExclamation, "Configuração"
        Exit Sub
    End If
    
    horaEscolhida = InputBox("Digite o horário de disparo no formato HH:MM (exemplo: 09:30 ou 14:00):", _
                             "Configurar Horário de Envio", ObterHoraSalva())
    If Trim(horaEscolhida) = "" Then Exit Sub
    
    On Error Resume Next
    horaValida = TimeValue(horaEscolhida)
    If Err.Number <> 0 Then
        MsgBox "Formato de hora inválido! Use o formato HH:MM (ex: 09:30).", vbCritical, "Erro de Horário"
        Exit Sub
    End If
    On Error GoTo 0
    
    wsP.Range(CEL_ATIVO).Value = "SIM"
    wsP.Range(CEL_DIA).Value = numDia
    wsP.Range(CEL_HORA).Value = Format(horaValida, "hh:nn")
    
    MsgBox "Configurações salvas com sucesso!" & vbCrLf & vbCrLf & _
           "Dia: " & NomeDoDia(numDia) & vbCrLf & _
           "Horário: " & Format(horaValida, "hh:nn") & vbCrLf & _
           "Envio automático: ATIVADO", vbInformation, "Agendamento Configurado"
End Sub

Public Sub Atualizar_Tabela_Regioes_Na_Aba_Parametros()
    Dim ws As Worksheet
    CriarAbaParametrosSeNaoExistir
    Set ws = ThisWorkbook.Worksheets(ABA_PARAM)
    
    ws.Range("A7").Value = "Região"
    ws.Range("B7").Value = "Estados (UF)"
    ws.Range("C7").Value = "Consultores em Cópia (CC)"
    
    ws.Range("A8").Value = "Sul"
    ws.Range("B8").Value = "RS, SC, PR"
    ws.Range("C8").Value = "consultor11@exemplo.com.br; consultor12@exemplo.com.br; consultor13@exemplo.com.br; consultor14@exemplo.com.br"
    
    ws.Range("A9").Value = "Sudeste SP"
    ws.Range("B9").Value = "SP"
    ws.Range("C9").Value = "consultor15@exemplo.com.br; consultor16@exemplo.com.br; consultor17@exemplo.com.br; consultor18@exemplo.com.br; consultor19@exemplo.com.br"
    
    ws.Range("A10").Value = "Sudeste MG"
    ws.Range("B10").Value = "MG"
    ws.Range("C10").Value = "consultor20@exemplo.com.br; consultor21@exemplo.com.br; consultor22@exemplo.com.br; consultor23@exemplo.com.br; consultor24@exemplo.com.br"
    
    ws.Range("A11").Value = "Sudeste RJ e ES"
    ws.Range("B11").Value = "RJ, ES"
    ws.Range("C11").Value = "consultor25@exemplo.com.br; consultor26@exemplo.com.br; consultor27@exemplo.com.br; consultor28@exemplo.com.br; consultor29@exemplo.com.br"
    
    ws.Range("A12").Value = "Nordeste"
    ws.Range("B12").Value = "BA, PE, CE, MA, PB, RN, AL, SE, PI"
    ws.Range("C12").Value = "consultor01@exemplo.com.br; consultor02@exemplo.com.br; consultor03@exemplo.com.br; consultor04@exemplo.com.br; consultor05@exemplo.com.br; consultor06@exemplo.com.br; consultor07@exemplo.com.br; consultor08@exemplo.com.br; consultor09@exemplo.com.br; consultor10@exemplo.com.br"
    
    ws.Range("A13").Value = "Centro Oeste"
    ws.Range("B13").Value = "DF, GO, MT, MS, AC, AP, AM, PA, RO, RR, TO"
    ws.Range("C13").Value = "consultor30@exemplo.com.br; consultor31@exemplo.com.br"
    
    ws.Range("A7:C7").Font.Bold = True
    ws.Columns("A:C").AutoFit
    
    MsgBox "Tabela de Regiões atualizada com sucesso na aba Parametros!", vbInformation, "Regiões Atualizadas"
End Sub

' ---------------------------------------------------------------------------
'  3. MACROS PRINCIPAIS ("EM ANDAMENTO")
' ---------------------------------------------------------------------------

Public Sub Revisar_Emails_Em_Andamento()
    Atualizar_Base_Em_Andamento
    ProcessarEmails False, ""
End Sub

Public Sub Enviar_Emails_Em_Andamento()
    Dim n As Long
    Atualizar_Base_Em_Andamento
    n = ContarPendentes()
    If n = 0 Then
        MsgBox "Nenhuma STA com peças 'Em andamento' pendente de envio." & vbCrLf & _
               "Se você já enviou hoje, use 'Limpar_Log_Para_Reenviar' para liberar novo envio.", vbInformation, "Coleta Em Andamento"
        Exit Sub
    End If
    If MsgBox("Deseja enviar agora " & n & " e-mail(s) de aviso de coleta para STAs 'Em Andamento'?", _
              vbYesNo + vbQuestion, "Aviso de Coleta - Em Andamento") <> vbYes Then Exit Sub
    ProcessarEmails True, ""
End Sub

Public Sub Limpar_Log_Para_Reenviar()
    Dim ws As Worksheet, ult As Long
    Set ws = ThisWorkbook.Worksheets(ABA_EMAILS)
    ult = UltimaLinha(ws)
    If ult >= 2 Then
        ws.Range(ws.Cells(2, COL_LOG), ws.Cells(ult, COL_LOG)).ClearContents
        MsgBox "Coluna de Log limpa com sucesso! Agora você pode reenviar.", vbInformation, "Log Limpo"
    End If
End Sub

Public Sub Envio_Automatico_Em_Andamento()
    Dim liberadas As Long
    silencioso = True
    On Error GoTo falha

    Application.DisplayAlerts = False
    Atualizar_Base_Em_Andamento
    liberadas = LiberarEnviosAntigos(DIAS_CICLO)
    Gravar "Início do envio automático (Coleta Em andamento) - STAs liberadas: " & liberadas
    ProcessarEmails True, ""
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    silencioso = False
    Exit Sub

falha:
    Gravar "ERRO " & Err.Number & ": " & Err.Description
    Application.DisplayAlerts = True
    silencioso = False
End Sub

Public Sub Enviar_Email_Teste_Em_Andamento()
    Dim destino As String
    destino = InputBox("Digite o e-mail de teste que receberá a mensagem:", "Teste - Em Andamento", "analista.rptg@exemplo.com.br")
    If Trim(destino) = "" Then Exit Sub
    Atualizar_Base_Em_Andamento
    ProcessarEmails True, destino
End Sub

Public Sub Atualizar_Base_Em_Andamento()
    Dim wsB As Worksheet, wsE As Worksheet
    Dim dic As Object, dicUF As Object, chave As String, cod As String, uf As String
    Dim dados As Variant, i As Long, ult As Long
    Dim cCod As Long, cQtd As Long, cVal As Long, cMes As Long, cAno As Long, cGrp As Long, cStatus As Long
    Dim mes As Long, ano As Long, periodos As Object, itens As Variant
    Dim texto As String, k As Variant
    Dim cMEI As Long, cUF As Long
    Dim regiaoNome As String, emailsCC As String
    Dim stGrp As String, stItem As String, ehValido As Boolean
    Dim valorTotal As Double

    Atualizar_Aging_Em_Andamento

    Set wsB = ThisWorkbook.Worksheets(ABA_BASE)
    Set wsE = ThisWorkbook.Worksheets(ABA_EMAILS)

    cCod = Coluna(wsB, "Codigo_STA")
    cQtd = Coluna(wsB, "QTD")
    cVal = Coluna(wsB, "Valor_Total")
    cMes = Coluna(wsB, "Mes")
    cAno = Coluna(wsB, "Ano")
    cGrp = Coluna(wsB, "Grupo_Status")
    cMEI = Coluna(wsB, "MEI")
    cUF = Coluna(wsB, "UF")
    
    cStatus = Coluna(wsB, "Status")
    If cStatus = 0 Then cStatus = Coluna(wsB, "Status_Chamado")
    If cStatus = 0 Then cStatus = Coluna(wsB, "Status_OS")
    If cStatus = 0 Then cStatus = Coluna(wsB, "Descricao_Status")
    If cStatus = 0 Then cStatus = Coluna(wsB, "Status_Processo")
    
    If cCod = 0 Or cQtd = 0 Or cVal = 0 Or cMes = 0 Or cAno = 0 Or (cGrp = 0 And cStatus = 0) Or cUF = 0 Then Exit Sub

    ult = wsB.Cells(wsB.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Set dic = CreateObject("Scripting.Dictionary")
    Set dicUF = CreateObject("Scripting.Dictionary")
    dados = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, wsB.UsedRange.Columns.Count)).Value

    For i = 1 To UBound(dados, 1)
        stGrp = ""
        stItem = ""
        
        If cGrp > 0 Then stGrp = UCase$(Trim$(CStr(dados(i, cGrp))))
        If cStatus > 0 Then stItem = UCase$(Trim$(CStr(dados(i, cStatus))))

        ehValido = False
        
        ' 1. Aceita se grupo ou status for EM ANDAMENTO
        If stGrp = STATUS_ALVO Or stItem = STATUS_ALVO Or stItem = "ANDAMENTO" Then
            ehValido = True
        End If
        
        ' 2. EXCLUSÃO ESTRITA: Se contiver TRANSPORTE em qualquer campo, descarta
        If InStr(1, stGrp, "TRANSPORTE") > 0 Or InStr(1, stItem, "TRANSPORTE") > 0 Then
            ehValido = False
        End If
        
        ' 3. Se houver status detalhado e for coletado / em transporte, descarta
        If cStatus > 0 And stItem <> "" Then
            If stItem = "EM TRANSPORTE" Or stItem = "COLETADO" Or InStr(1, stItem, "TRANSPORTE") > 0 Then
                ehValido = False
            End If
        End If

        If ehValido Then
            cod = ChaveSTA(dados(i, cCod))
            If cod <> "" Then
                uf = UCase$(Trim$(CStr(dados(i, cUF))))
                dicUF(cod) = uf
                
                ano = Val(dados(i, cAno))
                mes = Val(dados(i, cMes))
                chave = Format$(ano, "0000") & "-" & Format$(mes, "00")
                
                If Not dic.Exists(cod) Then
                    Set periodos = CreateObject("Scripting.Dictionary")
                    dic.Add cod, periodos
                End If
                Set periodos = dic(cod)
                
                If periodos.Exists(chave) Then
                    itens = periodos(chave)
                Else
                    itens = Array(CDbl(0), CDbl(0))
                End If
                itens = Array(itens(0) + Val(dados(i, cQtd)), itens(1) + Val(dados(i, cVal)))
                periodos(chave) = itens
            End If
        End If
    Next i

    ult = UltimaLinha(wsE)
    wsE.Cells(1, COL_DETALHE).Value = "Meses em aberto (Em andamento)"
    wsE.Cells(1, COL_REGRA).Value = "Status"
    wsE.Cells(1, COL_DEBITO).Value = "Valor Total (R$)"
    wsE.Cells(1, COL_REGIAO).Value = "Região"
    wsE.Cells(1, COL_CC).Value = "Consultores em Cópia (CC)"
    
    If ult >= 2 Then wsE.Range(wsE.Cells(2, COL_DETALHE), wsE.Cells(ult, COL_CC)).ClearContents

    For i = 2 To ult
        cod = ChaveSTA(wsE.Cells(i, 4).Value)
        texto = ""
        valorTotal = 0
        regiaoNome = ""
        emailsCC = ""
        
        If dic.Exists(cod) Then
            Set periodos = dic(cod)
            For Each k In OrdenarChaves(periodos.Keys)
                itens = periodos(k)
                texto = texto & IIf(texto = "", "", vbLf) & _
                        Mid$(CStr(k), 6, 2) & "/" & Left$(CStr(k), 4) & ": " & _
                        Format$(itens(0), "#,##0") & " peça(s) — R$ " & _
                        Format$(itens(1), "#,##0.00")
                valorTotal = valorTotal + itens(1)
            Next k
            
            If dicUF.Exists(cod) Then
                ObterRegiaoEConsultores dicUF(cod), regiaoNome, emailsCC
            End If
        End If
        
        wsE.Cells(i, COL_DETALHE).Value = texto
        wsE.Cells(i, COL_REGRA).Value = IIf(texto <> "", "Pendente Coleta", "")
        If valorTotal > 0 Then wsE.Cells(i, COL_DEBITO).Value = valorTotal
        wsE.Cells(i, COL_REGIAO).Value = regiaoNome
        wsE.Cells(i, COL_CC).Value = emailsCC
    Next i

    wsE.Columns(COL_DETALHE).ColumnWidth = 46
    wsE.Columns(COL_REGRA).ColumnWidth = 20
    wsE.Columns(COL_DEBITO).ColumnWidth = 20
    wsE.Columns(COL_REGIAO).ColumnWidth = 24
    wsE.Columns(COL_CC).ColumnWidth = 40
    wsE.Range(wsE.Cells(2, COL_DEBITO), wsE.Cells(ult, COL_DEBITO)).NumberFormat = "#,##0.00"
    Application.ScreenUpdating = True

    If silencioso Then
        Gravar "Atualização concluída para " & dic.Count & " STA(s) em andamento."
    End If
End Sub

Private Function ChaveSTA(ByVal valor As Variant) As String
    Dim s As String
    s = Trim$(UCase$(CStr(valor)))
    s = Replace(s, "STA", "")
    s = Trim$(s)
    If IsNumeric(s) And s <> "" Then
        ChaveSTA = CStr(CLng(Val(s)))
    Else
        ChaveSTA = s
    End If
End Function

Public Sub Atualizar_Aging_Em_Andamento()
    Dim ws As Worksheet
    Dim cGrp As Long, cStatus As Long, cDta As Long, cDias As Long, cFaixa As Long, cCham As Long
    Dim dados As Variant, dias As Variant, faixa As Variant
    Dim i As Long, ult As Long, n As Long, d As Double, hoje As Double
    Dim grp As String, stItem As String, ehValido As Boolean, calcAntes As XlCalculation

    Set ws = ThisWorkbook.Worksheets(ABA_BASE)
    cGrp = Coluna(ws, "Grupo_Status")
    cStatus = Coluna(ws, "Status")
    If cStatus = 0 Then cStatus = Coluna(ws, "Status_Chamado")
    If cStatus = 0 Then cStatus = Coluna(ws, "Status_OS")
    If cStatus = 0 Then cStatus = Coluna(ws, "Descricao_Status")
    If cStatus = 0 Then cStatus = Coluna(ws, "Status_Processo")
    
    cDta = Coluna(ws, "Dta_Encerramento")
    cDias = Coluna(ws, "Dias_Em_Aberto")
    cFaixa = Coluna(ws, "Faixa_Aging")
    cCham = Coluna(ws, "Chamado")
    
    If (cGrp = 0 And cStatus = 0) Or cDta = 0 Or cDias = 0 Or cFaixa = 0 Or cCham = 0 Then Exit Sub

    ult = ws.Cells(ws.Rows.Count, cCham).End(xlUp).Row
    If ult < 2 Then Exit Sub
    n = ult - 1

    calcAntes = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    dados = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value
    ReDim dias(1 To n, 1 To 1)
    ReDim faixa(1 To n, 1 To 1)
    hoje = CDbl(Date)

    For i = 1 To n
        dias(i, 1) = Empty
        faixa(i, 1) = Empty
        
        grp = ""
        stItem = ""
        If cGrp > 0 Then grp = UCase$(Trim$(CStr(dados(i, cGrp))))
        If cStatus > 0 Then stItem = UCase$(Trim$(CStr(dados(i, cStatus))))
        
        ehValido = False
        If grp = STATUS_ALVO Or stItem = STATUS_ALVO Or stItem = "ANDAMENTO" Then
            ehValido = True
        End If
        If InStr(1, grp, "TRANSPORTE") > 0 Or InStr(1, stItem, "TRANSPORTE") > 0 Then
            ehValido = False
        End If
        
        If ehValido Then
            If IsDate(dados(i, cDta)) Then
                d = hoje - CDbl(CDate(dados(i, cDta)))
                dias(i, 1) = d
                If d <= 30 Then
                    faixa(i, 1) = "0-30 dias"
                ElseIf d <= 60 Then
                    faixa(i, 1) = "31-60 dias"
                ElseIf d <= 90 Then
                    faixa(i, 1) = "61-90 dias"
                Else
                    faixa(i, 1) = "Mais de 90 dias"
                End If
            End If
        End If
    Next i

    ws.Cells(2, cDias).Resize(n, 1).Value = dias
    ws.Cells(2, cFaixa).Resize(n, 1).Value = faixa

    Application.Calculation = calcAntes
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
'  4. MAPEAMENTO DE REGIÕES E CONSULTORES (LÊ DA ABA PARAMETROS)
' ---------------------------------------------------------------------------
Private Sub ObterRegiaoEConsultores(ByVal uf As String, ByRef regiaoNome As String, ByRef emailsCC As String)
    Dim wsP As Worksheet
    Dim emailsPadrao As String
    Dim i As Long, ultP As Long, regTab As String
    
    uf = UCase$(Trim$(uf))
    
    Select Case uf
        ' SUL
        Case "RS", "SC", "PR"
            regiaoNome = "SUL"
            emailsPadrao = "consultor11@exemplo.com.br; consultor12@exemplo.com.br; consultor13@exemplo.com.br; consultor14@exemplo.com.br"
            
        ' SUDESTE SP
        Case "SP"
            regiaoNome = "SUDESTE SP"
            emailsPadrao = "consultor15@exemplo.com.br; consultor16@exemplo.com.br; consultor17@exemplo.com.br; consultor18@exemplo.com.br; consultor19@exemplo.com.br"
            
        ' SUDESTE MG
        Case "MG"
            regiaoNome = "SUDESTE MG"
            emailsPadrao = "consultor20@exemplo.com.br; consultor21@exemplo.com.br; consultor22@exemplo.com.br; consultor23@exemplo.com.br; consultor24@exemplo.com.br"
            
        ' SUDESTE RJ E ES
        Case "RJ", "ES"
            regiaoNome = "SUDESTE RJ/ES"
            emailsPadrao = "consultor25@exemplo.com.br; consultor26@exemplo.com.br; consultor27@exemplo.com.br; consultor28@exemplo.com.br; consultor29@exemplo.com.br"
            
        ' NORDESTE
        Case "BA", "PE", "CE", "MA", "PB", "RN", "AL", "SE", "PI"
            regiaoNome = "NORDESTE"
            emailsPadrao = "consultor01@exemplo.com.br; consultor02@exemplo.com.br; consultor03@exemplo.com.br; consultor04@exemplo.com.br; consultor05@exemplo.com.br; consultor06@exemplo.com.br; consultor07@exemplo.com.br; consultor08@exemplo.com.br; consultor09@exemplo.com.br; consultor10@exemplo.com.br"
            
        ' CENTRO-OESTE E NORTE
        Case "DF", "GO", "MT", "MS", "AC", "AP", "AM", "PA", "RO", "RR", "TO"
            regiaoNome = "CENTRO-OESTE / NORTE"
            emailsPadrao = "consultor30@exemplo.com.br; consultor31@exemplo.com.br"
            
        Case Else
            regiaoNome = "GERAL"
            emailsPadrao = ""
    End Select

    ' Tenta buscar se o usuário personalizou o e-mail na tabela da aba Parametros
    emailsCC = ""
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(ABA_PARAM)
    On Error GoTo 0
    
    If Not wsP Is Nothing Then
        ultP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
        For i = 1 To ultP
            regTab = UCase$(Trim$(CStr(wsP.Cells(i, 1).Value)))
            If regTab <> "" Then
                If regTab = UCase$(regiaoNome) Or InStr(1, regTab, UCase$(regiaoNome)) > 0 Or _
                   (regiaoNome = "SUDESTE RJ/ES" And (InStr(1, regTab, "RJ") > 0 Or InStr(1, regTab, "ES") > 0)) Then
                    If InStr(1, CStr(wsP.Cells(i, 3).Value), "@") > 0 Then
                        emailsCC = Trim$(CStr(wsP.Cells(i, 3).Value))
                    ElseIf InStr(1, CStr(wsP.Cells(i, 2).Value), "@") > 0 Then
                        emailsCC = Trim$(CStr(wsP.Cells(i, 2).Value))
                    End If
                    If emailsCC <> "" Then Exit For
                End If
            End If
        Next i
    End If
    
    If emailsCC = "" Then emailsCC = emailsPadrao
End Sub

' ---------------------------------------------------------------------------
'  5. MONTAGEM DO CORPO DO E-MAIL EM HTML
' ---------------------------------------------------------------------------
Private Function MontarHtmlColeta(ByVal codSTA As String, ByVal detalheMeses As String, _
                                  ByVal valorTotal As Double, ByVal temBanner As Boolean) As String
    Dim html As String, detalheHtml As String
    detalheHtml = TextoParaHtml(detalheMeses)
    
    html = "<div style=""font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222222; line-height: 1.5;"">"
    html = html & "<p>Prezado(a) Parceiro(a) <strong>STA " & codSTA & "</strong>,</p>"
    html = html & "<p>Identificamos em nosso sistema chamados de garantia (<strong>RPTG</strong>) com status " & _
                  "<strong style=""color: #0056b3;"">Em Andamento</strong> aguardando a finalização do processo de <strong>coleta das peças</strong>.</p>"
    
    html = html & "<div style=""background-color: #f4f6f9; border-left: 4px solid #0056b3; padding: 12px 16px; margin: 15px 0; border-radius: 4px;"">" & _
                  "<strong style=""color: #333333; font-size: 11pt;"">Resumo das peças pendentes de coleta:</strong><br><br>" & _
                  "<div style=""font-family: Consolas, monospace, sans-serif; font-size: 10.5pt; color: #2c3e50;"">" & _
                  detalheHtml & "</div>"
                  
    If valorTotal > 0 Then
        html = html & "<br><span style=""color: #0056b3; font-weight: bold;"">Valor Total das peças pendentes: R$ " & _
                      Format$(valorTotal, "#,##0.00") & "</span>"
    End If
    html = html & "</div>"
    
    html = html & "<p><strong style=""color: #c0392b;"">AÇÃO NECESSÁRIA — Favor dar andamento imediato na coleta:</strong></p>" & _
                  "<ol style=""margin-top: 5px; padding-left: 20px;"">" & _
                  "<li style=""margin-bottom: 6px;""><strong>Embalagem e Identificação:</strong> Certifique-se de que as peças estejam devidamente embaladas e identificadas com os respectivos números de OS/Chamado.</li>" & _
                  "<li style=""margin-bottom: 6px;""><strong>Documentação Fiscal:</strong> Mantenha a NF de remessa/devolução pronta para entrega ao transportador.</li>" & _
                  "<li style=""margin-bottom: 6px;""><strong>Divergência ou Atraso:</strong> Caso a coleta já tenha sido solicitada e a transportadora não tenha comparecido, <em>responda a este e-mail imediatamente</em> informando a situação.</li>" & _
                  "</ol>"
                  
    html = html & "<p>Contamos com sua colaboração para regularizarmos essas devoluções o quanto antes.</p>"
    
    html = html & "<p style=""margin-top: 25px; margin-bottom: 4px;"">Atenciosamente,</p>" & _
                  "<div style=""font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #1f1f1f; line-height: 1.4;"">" & _
                  "<strong style=""font-size: 11.5pt; color: #003366;"">Nome do Analista</strong><br>" & _
                  "<span style=""color: #333333;"">Analista de Qualidade - RPTG</span><br>" & _
                  "<span style=""color: #555555;"">São Paulo - Brasil</span><br>" & _
                  "<span>E-mail: <a href=""mailto:analista.rptg@exemplo.com.br"" style=""color: #0056b3; text-decoration: none;"">analista.rptg@exemplo.com.br</a></span>" & _
                  "</div>"
    
    If temBanner Then
        html = html & "<br><img src=""cid:bannerempresa"" width=""" & LARGURA_BANNER & """ style=""border:0;display:block;margin-top:10px;"">"
    End If
    
    html = html & "</div>"
    MontarHtmlColeta = "<html><body>" & html & "</body></html>"
End Function

' ---------------------------------------------------------------------------
'  6. DISPARO / REVISÃO DOS E-MAILS
' ---------------------------------------------------------------------------
Private Sub ProcessarEmails(ByVal enviar As Boolean, ByVal destinoTeste As String)
    Dim ws As Worksheet
    Dim olApp As Object, olMail As Object, olAnexo As Object
    Dim i As Long, ult As Long, n As Long, falhas As Long
    Dim para As String, codSTA As String, assunto As String
    Dim detalhe As String, regiao As String, ccConsultores As String, ccTotal As String
    Dim banner As String, temBanner As Boolean
    Dim valorTotal As Double

    Set ws = ThisWorkbook.Worksheets(ABA_EMAILS)
    ult = UltimaLinha(ws)
    banner = CaminhoBanner()
    temBanner = (banner <> "")

    If Not temBanner And Not silencioso Then
        If MsgBox("Não encontrei o arquivo " & ARQ_BANNER & " na pasta da planilha." & vbCrLf & _
                  "Continuar sem o banner na assinatura?", vbYesNo + vbExclamation, _
                  "Cobrança Em Andamento") <> vbYes Then Exit Sub
    End If

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        If silencioso Then
            Gravar "Erro: Outlook indisponível"
        Else
            MsgBox "Não foi possível abrir o Outlook.", vbCritical
        End If
        Exit Sub
    End If

    If Trim(CStr(ws.Cells(1, COL_LOG).Value)) = "" Then
        ws.Cells(1, COL_LOG).Value = "Último envio"
    End If

    Application.ScreenUpdating = False
    For i = 2 To ult
        If LinhaValida(ws, i, checarLog:=enviar) Then
            para = Trim(CStr(ws.Cells(i, 1).Value))
            codSTA = Trim(CStr(ws.Cells(i, 4).Value))
            detalhe = CStr(ws.Cells(i, COL_DETALHE).Value)
            valorTotal = Val(ws.Cells(i, COL_DEBITO).Value)
            regiao = Trim(CStr(ws.Cells(i, COL_REGIAO).Value))
            ccConsultores = Trim(CStr(ws.Cells(i, COL_CC).Value))
            
            If regiao = "" Then regiao = "GERAL"
            assunto = "[AVISO RPTG - " & regiao & "] Ação Necessária: Andamento na Coleta de Peças - STA " & codSTA
            
            ccTotal = COPIA_FIXA
            If ccConsultores <> "" Then ccTotal = ccTotal & "; " & ccConsultores
            
            If destinoTeste <> "" Then
                para = destinoTeste
                assunto = "[TESTE] " & assunto
            ElseIf PARA_FIXO <> "" Then
                para = para & "; " & PARA_FIXO
            End If

            On Error Resume Next
            Set olMail = olApp.CreateItem(0)
            With olMail
                .To = para
                If ccTotal <> "" Then .cc = ccTotal
                .Subject = assunto
                If temBanner Then
                    Set olAnexo = .Attachments.Add(banner, 1, 0)
                    olAnexo.PropertyAccessor.SetProperty _
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "bannerempresa"
                    Set olAnexo = Nothing
                End If
                .HTMLBody = MontarHtmlColeta(codSTA, detalhe, valorTotal, temBanner)
                If enviar Then
                    .Send
                Else
                    .Display
                End If
            End With
            
            If Err.Number <> 0 Then
                falhas = falhas + 1
                ws.Cells(i, COL_LOG).Value = "Erro: " & Err.Description
                Err.Clear
            Else
                n = n + 1
                If enviar And destinoTeste = "" Then ws.Cells(i, COL_LOG).Value = Now
            End If
            On Error GoTo 0
            Set olMail = Nothing

            If destinoTeste <> "" Then Exit For
            If Not enviar And MAX_REVISAR > 0 And n >= MAX_REVISAR Then Exit For
            If enviar And PAUSA_SEGUNDOS > 0 Then
                Application.Wait Now + TimeSerial(0, 0, 1)
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    If silencioso Then
        Gravar "E-mails enviados (Coleta Em Andamento): " & n & "; falhas: " & falhas
    Else
        MsgBox IIf(enviar, "E-mails enviados com sucesso: ", "E-mails abertos para revisão: ") & n & _
               IIf(falhas > 0, vbCrLf & "Falhas: " & falhas & " (veja a coluna I)", ""), vbInformation, "Coleta Em Andamento"
    End If
End Sub

' ---------------------------------------------------------------------------
'  7. FUNÇÕES AUXILIARES E PARÂMETROS
' ---------------------------------------------------------------------------
Private Function ObterDiaSalvo() As Integer
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_DIA).Value
    On Error GoTo 0
    If IsNumeric(v) And Val(v) >= 2 Then
        ObterDiaSalvo = CInt(v)
    Else
        ObterDiaSalvo = 2
    End If
End Function

Private Function ObterHoraSalva() As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_HORA).Text
    On Error GoTo 0
    If Trim(CStr(v)) <> "" Then
        ObterHoraSalva = Trim(CStr(v))
    Else
        ObterHoraSalva = "09:30"
    End If
End Function

Private Function NomeDoDia(ByVal dia As Integer) As String
    Select Case dia
        Case 2: NomeDoDia = "Segunda-feira"
        Case 3: NomeDoDia = "Terça-feira"
        Case 4: NomeDoDia = "Quarta-feira"
        Case 5: NomeDoDia = "Quinta-feira"
        Case 6: NomeDoDia = "Sexta-feira"
        Case 8: NomeDoDia = "Todos os dias úteis (Seg a Sex)"
        Case Else: NomeDoDia = "Não definido"
    End Select
End Function

Private Sub CriarAbaParametrosSeNaoExistir()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_PARAM)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ABA_PARAM
    End If
    If ws.Range("A2").Value = "" Then
        ws.Range("A2").Value = "Envio Automático Ativo:"
        ws.Range(CEL_ATIVO).Value = "SIM"
        ws.Range("A3").Value = "Dia da Semana (2=Seg..6=Sex, 8=Diário):"
        ws.Range(CEL_DIA).Value = 2
        ws.Range("A4").Value = "Horário de Disparo:"
        ws.Range(CEL_HORA).Value = "09:30"
        ws.Columns("A:B").AutoFit
    End If
End Sub

Private Function LinhaValida(ws As Worksheet, ByVal i As Long, Optional ByVal checarLog As Boolean = True) As Boolean
    If InStr(1, CStr(ws.Cells(i, 1).Value), "@") = 0 Then Exit Function
    If checarLog And PULAR_JA_ENVIADOS And Trim(CStr(ws.Cells(i, COL_LOG).Value)) <> "" Then Exit Function
    If Trim(CStr(ws.Cells(i, COL_DETALHE).Value)) = "" Then Exit Function
    LinhaValida = True
End Function

Private Function ContarPendentes() As Long
    Dim ws As Worksheet, i As Long, ult As Long, n As Long
    Set ws = ThisWorkbook.Worksheets(ABA_EMAILS)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If LinhaValida(ws, i, checarLog:=True) Then n = n + 1
    Next i
    ContarPendentes = n
End Function

Private Function UltimaLinha(ws As Worksheet) As Long
    Dim a As Long, d As Long
    a = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    d = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    UltimaLinha = IIf(a > d, a, d)
End Function

Private Function Coluna(ws As Worksheet, ByVal titulo As String) As Long
    Dim j As Long
    For j = 1 To ws.UsedRange.Columns.Count
        If Trim$(CStr(ws.Cells(1, j).Value)) = titulo Then
            Coluna = j
            Exit Function
        End If
    Next j
End Function

Private Function OrdenarChaves(ByVal chaves As Variant) As Variant
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If CStr(chaves(j)) < CStr(chaves(i)) Then
                tmp = chaves(i)
                chaves(i) = chaves(j)
                chaves(j) = tmp
            End If
        Next j
    Next i
    OrdenarChaves = chaves
End Function

Private Function LiberarEnviosAntigos(ByVal dias As Long) As Long
    Dim ws As Worksheet, i As Long, ult As Long, n As Long, v As Variant
    Set ws = ThisWorkbook.Worksheets(ABA_EMAILS)
    ult = UltimaLinha(ws)
    For i = 2 To ult
        v = ws.Cells(i, COL_LOG).Value
        If Trim(CStr(v)) <> "" Then
            If IsDate(v) Then
                If Now - CDate(v) >= dias Then
                    ws.Cells(i, COL_LOG).ClearContents
                    n = n + 1
                End If
            Else
                ws.Cells(i, COL_LOG).ClearContents
                n = n + 1
            End If
        End If
    Next i
    LiberarEnviosAntigos = n
End Function

Private Sub Gravar(ByVal texto As String)
    Dim pasta As String, arq As String, fnum As Integer
    On Error Resume Next
    pasta = PastaLocalDaPlanilha()
    If pasta = "" Then pasta = Environ$("TEMP")
    arq = pasta & Application.PathSeparator & ARQ_LOG
    fnum = FreeFile
    Open arq For Append As #fnum
    Print #fnum, Format$(Now, "dd/mm/yyyy hh:nn") & " - " & texto
    Close #fnum
    On Error GoTo 0
End Sub

Public Sub Escolher_Banner()
    Dim dlg As Object, caminho As String
    Set dlg = Application.FileDialog(3)
    dlg.Title = "Selecione a imagem do banner da assinatura"
    dlg.AllowMultiSelect = False
    dlg.Filters.Clear
    dlg.Filters.Add "Imagens", "*.png; *.jpg; *.jpeg; *.gif"
    If dlg.Show <> -1 Then Exit Sub
    caminho = dlg.SelectedItems(1)
    GravarCaminhoBanner caminho
    MsgBox "Banner definido:" & vbCrLf & caminho, vbInformation, "Banner Definido"
End Sub

Private Sub GravarCaminhoBanner(ByVal caminho As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_BANNER).Value = caminho
    On Error GoTo 0
End Sub

Private Function CaminhoGravado() As String
    On Error Resume Next
    CaminhoGravado = Trim(CStr(ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_BANNER).Value))
    On Error GoTo 0
End Function

Private Function PastaLocalDaPlanilha() As String
    Dim p As String, resto As String, i As Long, base As Variant
    p = ThisWorkbook.Path
    If p = "" Then Exit Function
    If LCase(Left$(p, 4)) <> "http" Then
        PastaLocalDaPlanilha = p
        Exit Function
    End If
    resto = Replace(p, "/", Application.PathSeparator)
    i = InStr(1, resto, Application.PathSeparator & "Documents" & Application.PathSeparator, vbTextCompare)
    If i = 0 Then i = InStr(1, resto, Application.PathSeparator & "Documentos" & Application.PathSeparator, vbTextCompare)
    If i = 0 Then Exit Function
    resto = Mid$(resto, InStr(i + 1, resto, Application.PathSeparator))
    For Each base In Array(Environ$("OneDriveCommercial"), Environ$("OneDriveConsumer"), Environ$("OneDrive"))
        If CStr(base) <> "" Then
            If Dir(CStr(base) & resto, vbDirectory) <> "" Then
                PastaLocalDaPlanilha = CStr(base) & resto
                Exit Function
            End If
        End If
    Next base
End Function

Private Function CaminhoBanner() As String
    Dim caminho As String, pasta As String
    If CAMINHO_BANNER <> "" Then
        If Dir(CAMINHO_BANNER) <> "" Then CaminhoBanner = CAMINHO_BANNER
        Exit Function
    End If
    caminho = CaminhoGravado()
    If caminho <> "" Then
        If Dir(caminho) <> "" Then
            CaminhoBanner = caminho
            Exit Function
        End If
    End If
    pasta = PastaLocalDaPlanilha()
    If pasta <> "" Then
        caminho = pasta & Application.PathSeparator & ARQ_BANNER
        If Dir(caminho) <> "" Then
            CaminhoBanner = caminho
            GravarCaminhoBanner caminho
        End If
    End If
End Function

Private Function TextoParaHtml(ByVal texto As String) As String
    Dim s As String
    s = Replace(texto, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "<br>")
    TextoParaHtml = s
End Function