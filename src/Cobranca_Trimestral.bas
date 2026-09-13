Attribute VB_Name = "Cobranca_Trimestral"
'=====================================================================
' COBRANCA DE DEVOLUCOES (trimestre) E PLANILHA DE DEBITO
'
' Regra do prazo: peca do mes M (colunas Mes/Ano da base) tem prazo ate o
' dia 15 do 3o mes seguinte (marco -> 15/06, abril -> 15/07, maio -> 15/08).
'
' 1) Comeco do mes  -> Revisar_Cobranca_STAs / Enviar_Cobranca_STAs
'    E-mail por STA com as pecas nao devolvidas e o prazo de cada uma.
'    Assunto personalizado por Regiao e Consultores em Copia (CC).
' 2) Fim do mes      -> Planilha_Debito_Financeiro
'    Gera a planilha das pecas que passaram do prazo e o e-mail para o
'    Financeiro com o arquivo anexado.
'
' Nao cobra: STA MEI e STAs da regiao Norte (excluidas dos dois envios).
'
' Cadastro na aba Parametros:
'   P6  = caminho da imagem do banner da assinatura (opcional)
'   P20 = e-mail do Financeiro (pode ter varios separados por ;)
'   P21 = pasta onde salvar a planilha de debito (vazio = pasta da planilha)
'
' Todo o texto do e-mail e HTML com entidades (&ccedil; etc.), por isso
' o modulo nao tem nenhum acento no codigo e nao quebra por encoding.
'=====================================================================
Option Explicit

Private Const CB_ABA_BASE   As String = "Base_RPTG"
Private Const CB_NOME_TAB   As String = "tblBase"
Private Const CB_ABA_STAS   As String = "STAs"
Private Const CB_ABA_PARAM  As String = "Parametros"
Private Const CB_CEL_FINANCEIRO  As String = "P20"
Private Const CB_CEL_PASTA  As String = "P21"
Private Const CB_DIA_LIMITE As Long = 15   ' dia do mes em que vence o prazo
Private Const CB_MESES      As Long = 3    ' mes da peca + 3 meses = mes do prazo
Private Const CB_SO_NO_PRAZO As Boolean = True ' e-mail ao STA cobra somente os meses ainda
                                              ' no prazo; os vencidos vao para o debito
Private Const CB_MOSTRAR_VALORES As Boolean = False  ' True mostra R$ no e-mail da STA
Private Const CB_ARQ_LOG    As String = "Log_Cobranca_RPTG.txt"
Private Const CB_ABA_CONFIG As String = "Config_Debito"
Private Const CB_ARQ_BANNER As String = "Assinatura_Empresa.png"
Private Const CB_CEL_BANNER As String = "P6"   ' caminho do banner (mesma celula do outro modulo)
Private Const CB_LARG_BANNER As Long = 500
Private Const CB_DIA_DEBITO As Long = 28   ' dia do mes do envio automatico do debito
Private Const CB_HORA_DEBITO As String = "15:00"

' data usada como "hoje" (permite gerar o debito de um mes escolhido)
Private CB_DataRef As Date
Private silenciosoDebito As Boolean
' meses de encerramento aceitos no debito (chave "aaaa-mm"); vazio = todos os vencidos
Private CB_MesesFiltro As Object

' True se o mes de encerramento entra no debito
Private Function CB_MesAceito(ByVal dtaEnc As Date) As Boolean
    If CB_MesesFiltro Is Nothing Then
        CB_MesAceito = True
    ElseIf CB_MesesFiltro.Count = 0 Then
        CB_MesAceito = True
    Else
        CB_MesAceito = CB_MesesFiltro.Exists(Format$(dtaEnc, "yyyy-mm"))
    End If
End Function

' True quando o usuario escolheu meses especificos (entram mesmo sem prazo vencido)
Private Function CB_MesEscolhido() As Boolean
    If CB_MesesFiltro Is Nothing Then Exit Function
    CB_MesEscolhido = (CB_MesesFiltro.Count > 0)
End Function

' le "04/2026;05/2026" ou "4,5" e devolve o dicionario de meses aceitos
Private Function CB_LerMeses(ByVal Txt As String, ByVal anoPadrao As Long) As Object
    Dim d As Object, partes As Variant, i As Long, p As String
    Dim mes As Long, ano As Long

    Set d = CreateObject("Scripting.Dictionary")
    Txt = Trim$(Replace(Replace(Replace(Txt, ";", ","), " ", ""), "-", ","))
    If Txt = "" Then
        Set CB_LerMeses = d
        Exit Function
    End If

    partes = Split(Txt, ",")
    For i = LBound(partes) To UBound(partes)
        p = CStr(partes(i))
        If p <> "" Then
            ano = anoPadrao
            If InStr(p, "/") > 0 Then
                mes = Val(Split(p, "/")(0))
                ano = Val(Split(p, "/")(1))
                If ano < 100 And ano > 0 Then ano = 2000 + ano
            Else
                mes = Val(p)
            End If
            If mes >= 1 And mes <= 12 And ano > 1900 Then
                d(Format$(ano, "0000") & "-" & Format$(mes, "00")) = True
            End If
        End If
    Next i
    Set CB_LerMeses = d
End Function

Private Function CB_Hoje() As Date
    If CB_DataRef > 0 Then CB_Hoje = CB_DataRef Else CB_Hoje = Date
End Function

'---------------------------------------------------------------------
' ROTINAS PUBLICAS
'---------------------------------------------------------------------
' comeco do mes: abre os e-mails para voce revisar antes de enviar
Public Sub Revisar_Cobranca_STAs()
    CB_Cobranca False
End Sub

' comeco do mes: envia direto
Public Sub Enviar_Cobranca_STAs()
    If MsgBox("Enviar os e-mails de cobranca agora?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    CB_Cobranca True
End Sub

' para o Agendador de Tarefas (sem nenhuma janela)
Public Sub Cobranca_Automatica()
    On Error Resume Next
    CB_Cobranca True
End Sub

' cobranca semanal (mesma regra) - use no Agendador toda semana
Public Sub Cobranca_Semanal_Automatica()
    On Error Resume Next
    CB_Cobranca True
End Sub

' fim do mes: gera a planilha e abre o e-mail para o Financeiro
Public Sub Planilha_Debito_Financeiro()
    CB_Debito False
End Sub

' fim do mes, para o Agendador: gera e envia sozinho
Public Sub Planilha_Debito_Financeiro_Automatico()
    On Error Resume Next
    CB_Debito True
End Sub

' fim do mes, escolhendo o mes/dia do corte: pergunta a data e gera a planilha
Public Sub Debito_Financeiro_Escolher_Data()
    Dim resp As String, dt As Date, env As VbMsgBoxResult

    resp = InputBox("Data de corte (dd/mm/aaaa)." & vbCrLf & vbCrLf & _
                    "Entram no debito as pecas cujo prazo (dia " & CB_DIA_LIMITE & _
                    ") venceu ate essa data." & vbCrLf & _
                    "O arquivo sai como Debito_RPTG_" & Format$(Date, "yyyy_mm") & ".xlsx", _
                    "Planilha de debito", Format$(Date, "dd/mm/yyyy"))
    If Trim$(resp) = "" Then Exit Sub
    If Not IsDate(resp) Then
        MsgBox "Data invalida. Use o formato dd/mm/aaaa.", vbExclamation
        Exit Sub
    End If
    dt = CDate(resp)

    resp = InputBox("Quais meses de encerramento entram no debito?" & vbCrLf & vbCrLf & _
                    "Ex.: 04/2026;05/2026   (ou apenas 4;5)" & vbCrLf & _
                    "Deixe em branco para incluir TODOS os meses com prazo vencido.", _
                    "Planilha de debito - meses", "")
    Set CB_MesesFiltro = CB_LerMeses(resp, Year(dt))

    env = MsgBox("Enviar o e-mail para o Financeiro automaticamente?" & vbCrLf & vbCrLf & _
                 "Sim = gera e envia" & vbCrLf & "Nao = gera e abre o e-mail para voce revisar", _
                 vbYesNoCancel + vbQuestion, "Planilha de debito - " & Format$(dt, "mm/yyyy"))
    If env = vbCancel Then
        Set CB_MesesFiltro = Nothing
        Exit Sub
    End If

    CB_DataRef = dt
    If env = vbYes Then
        CB_Debito True
    Else
        CB_Debito False
    End If
    CB_DataRef = 0
    Set CB_MesesFiltro = Nothing
End Sub

' envio automatico mensal (sem Agendador de Tarefas): a propria planilha
' dispara no dia CB_DIA_DEBITO, no horario CB_HORA_DEBITO, com ela aberta.
Public Sub Ativar_Debito_Mensal()
    Dim quando As Date
    quando = CB_ProximoDebito()
    On Error Resume Next
    Application.OnTime CB_ProximoDebito(), "Debito_Mensal_Automatico", , False
    On Error GoTo 0
    Application.OnTime quando, "Debito_Mensal_Automatico"
    CB_Log "Debito mensal agendado para " & Format$(quando, "dd/mm/yyyy hh:nn")
    If Not silenciosoDebito Then _
        MsgBox "Envio do debito agendado para " & Format$(quando, "dd/mm/yyyy hh:nn") & _
               "." & vbCrLf & "A planilha precisa estar aberta nesse horario.", vbInformation
End Sub

Public Sub Cancelar_Debito_Mensal()
    On Error Resume Next
    Application.OnTime CB_ProximoDebito(), "Debito_Mensal_Automatico", , False
    On Error GoTo 0
    MsgBox "Envio automatico do debito cancelado.", vbInformation
End Sub

' chamada pelo agendamento interno: gera, envia e reagenda o mes seguinte
Public Sub Debito_Mensal_Automatico()
    Dim ws As Worksheet, d As Object, linhas As Object, lista As String
    Dim k As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CB_ABA_CONFIG)
    On Error GoTo 0

    If ws Is Nothing Then
        CB_Log "Debito automatico nao enviado: aba " & CB_ABA_CONFIG & " nao existe."
    Else
        CB_LerMarcas ws, d, linhas, lista
        If d.Count = 0 Then
            CB_Log "Debito automatico nao enviado: nenhum mes marcado com X."
        Else
            On Error Resume Next
            CB_DataRef = Date
            Set CB_MesesFiltro = d
            CB_Debito True
            CB_DataRef = 0
            Set CB_MesesFiltro = Nothing
            For Each k In linhas.Keys
                ws.Cells(linhas(k), 7).Value = Format$(Date, "dd/mm/yyyy")
                ws.Cells(linhas(k), 6).ClearContents
            Next k
            CB_Log "Debito automatico enviado. Meses: " & lista
            On Error GoTo 0
        End If
    End If

    silenciosoDebito = True
    Ativar_Debito_Mensal
    silenciosoDebito = False
End Sub

' proximo dia CB_DIA_DEBITO as CB_HORA_DEBITO
Private Function CB_ProximoDebito() As Date
    Dim ini As Date, d As Date, dia As Long, i As Long

    ini = DateSerial(Year(Date), Month(Date), 1)
    For i = 0 To 1
        dia = CB_DIA_DEBITO
        If dia > Day(DateSerial(Year(ini), Month(ini) + 1, 0)) Then
            dia = Day(DateSerial(Year(ini), Month(ini) + 1, 0))
        End If
        d = DateSerial(Year(ini), Month(ini), dia) + TimeValue(CB_HORA_DEBITO)
        If d > Now Then Exit For
        ini = DateSerial(Year(ini), Month(ini) + 1, 1)
    Next i
    CB_ProximoDebito = d
End Function

' cadastra o e-mail do Financeiro
Public Sub Cadastrar_Email_Financeiro()
    Dim ws As Worksheet, s As String
    Set ws = CB_Param()
    If ws Is Nothing Then Exit Sub
    s = InputBox("E-mail do Financeiro (separe varios com ;):", "Planilha de debito", _
                 CStr(ws.Range(CB_CEL_FINANCEIRO).Value))
    If Trim$(s) = "" Then Exit Sub
    ws.Range(CB_CEL_FINANCEIRO).Value = s
    MsgBox "Gravado.", vbInformation
End Sub

' escolhe a pasta onde a planilha de debito sera salva
Public Sub Escolher_Pasta_Debito()
    Dim fd As Object, ws As Worksheet
    Set ws = CB_Param()
    If ws Is Nothing Then Exit Sub
    Set fd = Application.FileDialog(4)
    fd.Title = "Pasta onde salvar a planilha de debito"
    If fd.Show <> -1 Then Exit Sub
    ws.Range(CB_CEL_PASTA).Value = fd.SelectedItems(1)
    MsgBox "Gravado:" & vbCrLf & fd.SelectedItems(1), vbInformation
End Sub

' mostra quantas pecas estao dentro do prazo e quantas vencidas
Public Sub Conferir_Prazos()
    Dim dados As Object, k As Variant
    Dim nDentro As Long, nVenc As Long, nSta As Long
    Set dados = CB_Pendencias()
    For Each k In dados.Keys
        nSta = nSta + 1
        nDentro = nDentro + dados(k)("dentro").Count
        nVenc = nVenc + dados(k)("vencidas").Count
    Next k
    MsgBox "STAs com pendencia (fora MEI/Norte): " & nSta & vbCrLf & _
           "Pecas dentro do prazo: " & nDentro & vbCrLf & _
           "Pecas vencidas (entram no debito): " & nVenc, vbInformation, "Conferir prazos"
End Sub

'=====================================================================
' ABA Config_Debito: escolher os meses pela planilha
'=====================================================================
Public Sub Montar_Config_Debito()
    Dim dados As Object, ws As Worksheet, k As Variant, info As Object
    Dim it As Variant, i As Long, lin As Long
    Dim mes As Object, marca As Object, envio As Object
    Dim chaves As Variant, ch As String, tmp As Variant, j As Long
    Dim qtd As Double, vTot As Double, prazo As Date
    Dim col As Object

    Set marca = CreateObject("Scripting.Dictionary")
    Set envio = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CB_ABA_CONFIG)
    On Error GoTo 0

    If Not ws Is Nothing Then
        lin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lin
            ch = CB_ChaveMes(CB_Txt(ws.Cells(i, 1).Value))
            If ch <> "" Then
                marca(ch) = UCase$(CB_Txt(ws.Cells(i, 6).Value))
                envio(ch) = CB_Txt(ws.Cells(i, 7).Value)
            End If
        Next i
    Else
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = CB_ABA_CONFIG
    End If

    Set CB_MesesFiltro = Nothing
    Set dados = CB_Pendencias()
    Set mes = CreateObject("Scripting.Dictionary")

    For Each k In dados.Keys
        Set info = dados(k)
        For j = 1 To 2
            If j = 1 Then Set col = info("dentro") Else Set col = info("vencidas")
            For i = 1 To col.Count
                it = col(i)
                ch = Format$(it(4), "yyyy-mm")
                If Not mes.Exists(ch) Then mes(ch) = Array(0#, 0#, CDate(it(5)))
                mes(ch) = Array(mes(ch)(0) + it(3), mes(ch)(1) + it(8), CDate(it(5)))
            Next i
        Next j
    Next k

    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Mes referencia", "Prazo", "Situacao", _
        "Pecas", "Valor", "Enviar? (X)", "Enviado em")
    ws.Range("A1:G1").Font.Bold = True

    chaves = mes.Keys
    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If chaves(j) < chaves(i) Then
                tmp = chaves(i): chaves(i) = chaves(j): chaves(j) = tmp
            End If
        Next j
    Next i

    lin = 2
    For i = LBound(chaves) To UBound(chaves)
        ch = CStr(chaves(i))
        qtd = mes(ch)(0): vTot = mes(ch)(1): prazo = mes(ch)(2)
        ws.Cells(lin, 1).Value = Mid$(ch, 6, 2) & "/" & Left$(ch, 4)
        ws.Cells(lin, 2).Value = prazo
        ws.Cells(lin, 3).Value = IIf(Date > prazo, "Vencido", "No prazo")
        ws.Cells(lin, 4).Value = qtd
        ws.Cells(lin, 5).Value = vTot
        If envio.Exists(ch) Then ws.Cells(lin, 7).Value = envio(ch)
        If marca.Exists(ch) Then ws.Cells(lin, 6).Value = marca(ch)
        lin = lin + 1
    Next i

    ws.Range("B2:B" & lin).NumberFormat = "dd/mm/yyyy"
    ws.Range("E2:E" & lin).NumberFormat = "#,##0.00"
    ws.Columns("A:G").AutoFit
    ws.Activate

    MsgBox "Marque com X na coluna 'Enviar? (X)' os meses que vao para o Financeiro" & vbCrLf & _
           "e rode a macro Debito_Financeiro_Pela_Planilha." & vbCrLf & vbCrLf & _
           "Nenhum mes vem marcado: sem X o mes nao entra no debito." & vbCrLf & _
           "Nos meses que voce ja enviou, escreva a data na coluna 'Enviado em'.", _
           vbInformation, "Config_Debito"
End Sub

Public Sub Debito_Financeiro_Pela_Planilha()
    Dim ws As Worksheet, i As Long, ult As Long, ch As String
    Dim d As Object, lista As String, resp As String, dt As Date
    Dim env As VbMsgBoxResult, linhas As Object, k As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CB_ABA_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Rode primeiro a macro Montar_Config_Debito.", vbExclamation
        Exit Sub
    End If

    CB_LerMarcas ws, d, linhas, lista

    If d.Count = 0 Then
        MsgBox "Nenhum mes marcado com X na aba " & CB_ABA_CONFIG & ".", vbExclamation
        Exit Sub
    End If

    For Each k In linhas.Keys
        If CB_Txt(ws.Cells(linhas(k), 7).Value) <> "" Then
            If MsgBox("O mes " & Mid$(CStr(k), 6, 2) & "/" & Left$(CStr(k), 4) & _
                      " ja foi enviado em " & CB_Txt(ws.Cells(linhas(k), 7).Value) & "." & vbCrLf & _
                      "Enviar de novo?", vbYesNo + vbExclamation, "Debito do Financeiro") <> vbYes Then Exit Sub
        End If
    Next k

    resp = InputBox("Data de corte (dd/mm/aaaa) que vai no nome do arquivo e no e-mail.", _
                    "Debito do Financeiro", Format$(Date, "dd/mm/yyyy"))
    If Trim$(resp) = "" Then Exit Sub
    If Not IsDate(resp) Then
        MsgBox "Data invalida. Use o formato dd/mm/aaaa.", vbExclamation
        Exit Sub
    End If
    dt = CDate(resp)

    env = MsgBox("Meses marcados: " & lista & vbCrLf & vbCrLf & _
                 "Sim = gera e envia para o Financeiro" & vbCrLf & _
                 "Nao = gera e abre o e-mail para revisar", _
                 vbYesNoCancel + vbQuestion, "Debito do Financeiro")
    If env = vbCancel Then Exit Sub

    Set CB_MesesFiltro = d
    CB_DataRef = dt
    CB_Debito CBool(env = vbYes)
    CB_DataRef = 0
    Set CB_MesesFiltro = Nothing

    For Each k In linhas.Keys
        i = linhas(k)
        ws.Cells(i, 7).Value = Format$(dt, "dd/mm/yyyy")
        ws.Cells(i, 6).ClearContents
    Next k
End Sub

Private Sub CB_LerMarcas(ws As Worksheet, d As Object, linhas As Object, lista As String)
    Dim i As Long, ult As Long, ch As String

    Set d = CreateObject("Scripting.Dictionary")
    Set linhas = CreateObject("Scripting.Dictionary")
    lista = ""
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        If UCase$(CB_Txt(ws.Cells(i, 6).Value)) = "X" Then
            ch = CB_ChaveMes(CB_Txt(ws.Cells(i, 1).Value))
            If ch <> "" Then
                d(ch) = True
                linhas(ch) = i
                lista = lista & Mid$(ch, 6, 2) & "/" & Left$(ch, 4) & "  "
            End If
        End If
    Next i
End Sub

Private Function CB_ChaveMes(ByVal Txt As String) As String
    Dim mes As Long, ano As Long, p As Variant
    Txt = Trim$(Txt)
    If Txt = "" Then Exit Function
    If InStr(Txt, "/") > 0 Then
        p = Split(Txt, "/")
        mes = Val(p(0)): ano = Val(p(UBound(p)))
    ElseIf InStr(Txt, "-") > 0 Then
        p = Split(Txt, "-")
        ano = Val(p(0)): mes = Val(p(1))
    Else
        Exit Function
    End If
    If ano < 100 And ano > 0 Then ano = 2000 + ano
    If mes >= 1 And mes <= 12 And ano > 1900 Then _
        CB_ChaveMes = Format$(ano, "0000") & "-" & Format$(mes, "00")
End Function

'=====================================================================
' 1) COBRANCA POR STA (COM REGIAO NO ASSUNTO E CC CONSULTORES)
'=====================================================================
Private Sub CB_Cobranca(ByVal enviar As Boolean)
    Dim dados As Object, k As Variant
    Dim ol As Object, mail As Object
    Dim nMails As Long, msg As String
    Dim info As Object
    Dim regiaoNome As String, emailsConsultores As String

    Set dados = CB_Pendencias()
    If dados.Count = 0 Then
        If Not enviar Then MsgBox "Nenhuma pendencia a cobrar.", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        If Not enviar Then MsgBox "Outlook nao encontrado.", vbExclamation
        Exit Sub
    End If

    For Each k In dados.Keys
        Set info = dados(k)
        If info("email") <> "" And CB_QtdCobravel(info) > 0 Then
            ' Mapeia a regiao correta e os consultores em CC
            CB_ObterRegiaoEConsultores CStr(info("uf")), CStr(info("regiao")), regiaoNome, emailsConsultores
            
            Set mail = ol.CreateItem(0)
            mail.To = info("email")
            mail.cc = CB_Copias(emailsConsultores)
            mail.Subject = "[COBRANCA RPTG - " & regiaoNome & "] Devolucao de pecas pendentes - STA " & info("codigo") & " - " & info("nome")
            CB_AnexarBanner mail
            mail.HTMLBody = CB_Corpo(info)
            If enviar Then
                mail.Send
            Else
                mail.Display
            End If
            nMails = nMails + 1
            Set mail = Nothing
        End If
    Next k

    msg = "Cobranca " & IIf(enviar, "enviada", "aberta para revisao") & ": " & _
          nMails & " STA(s)."
    CB_Log msg
    If Not enviar Then MsgBox msg, vbInformation, "Cobranca de devolucoes"
End Sub

Private Function CB_QtdCobravel(ByVal info As Object) As Long
    If CB_SO_NO_PRAZO Then
        CB_QtdCobravel = info("dentro").Count
    Else
        CB_QtdCobravel = info("dentro").Count + info("vencidas").Count
    End If
End Function

Private Function CB_Corpo(ByVal info As Object) As String
    Dim h As String, i As Long, lin As Variant
    Dim saud As String

    If Hour(Now) < 12 Then saud = "Bom Dia" Else saud = "Boa Tarde"

    h = "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252""></head>" & _
        "<body style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">"
    h = h & "<p>" & saud & " !</p>"
    h = h & "<p>Segue a rela&ccedil;&atilde;o das pe&ccedil;as de RPTG <b>ainda n&atilde;o devolvidas</b> " & _
            "pela STA <b>" & CB_Html(CStr(info("nome"))) & "</b>.</p>"
    h = h & "<p>Cada pe&ccedil;a tem prazo de devolu&ccedil;&atilde;o at&eacute; o <b>dia " & CB_DIA_LIMITE & _
            "</b> conforme a coluna <i>Prazo</i>. " & _
            "Passado o prazo, a pe&ccedil;a &eacute; encaminhada para <b>d&eacute;bito</b>.</p>"
    If CB_SO_NO_PRAZO Then
        h = h & "<p>Esta rela&ccedil;&atilde;o traz apenas as pe&ccedil;as <b>ainda dentro do prazo</b>. " & _
                "As pe&ccedil;as com prazo j&aacute; vencido foram encaminhadas para <b>d&eacute;bito</b>.</p>"
        h = h & CB_Tabela("Pe&ccedil;as dentro do prazo", info("dentro"))
    Else
        If info("vencidas").Count > 0 Then
            h = h & "<p style=""color:#C00000""><b>Aten&ccedil;&atilde;o:</b> " & info("vencidas").Count & _
                    " pe&ccedil;a(s) j&aacute; est&atilde;o com o prazo vencido.</p>"
        End If
        h = h & CB_Tabela("Pe&ccedil;as com prazo vencido", info("vencidas"))
        h = h & CB_Tabela("Pe&ccedil;as dentro do prazo", info("dentro"))
    End If

    h = h & "<p>Ao enviar as pe&ccedil;as, encaminhe a nota fiscal de devolu&ccedil;&atilde;o para valida&ccedil;&atilde;o.</p>"
    h = h & CB_Assinatura()
    h = h & "</body></html>"
    CB_Corpo = h
End Function

Private Function CB_Assinatura() As String
    Dim h As String
    h = "<p>Atenciosamente,<br>" & _
        "Nome do Analista<br>" & _
        "Analista de Qualidade - RPTG<br>" & _
        "S&atilde;o Paulo - Brasil<br>" & _
        "E-mail: analista.rptg@exemplo.com.br</p>"
    If CB_CaminhoBanner() <> "" Then
        h = h & "<br><img src=""cid:bannerempresa"" width=""" & CB_LARG_BANNER & _
                """ style=""border:0;display:block;"">"
    End If
    CB_Assinatura = h
End Function

Private Sub CB_AnexarBanner(ByVal mail As Object)
    Dim caminho As String, anexo As Object
    caminho = CB_CaminhoBanner()
    If caminho = "" Then Exit Sub
    On Error Resume Next
    Set anexo = mail.Attachments.Add(caminho, 1, 0)
    anexo.PropertyAccessor.SetProperty _
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "bannerempresa"
    On Error GoTo 0
    Set anexo = Nothing
End Sub

Private Function CB_CaminhoBanner() As String
    Dim caminho As String, pasta As String
    On Error Resume Next
    caminho = Trim$(CB_Txt(ThisWorkbook.Worksheets(CB_ABA_PARAM).Range(CB_CEL_BANNER).Value))
    On Error GoTo 0
    If caminho <> "" Then
        If Dir(caminho) <> "" Then
            CB_CaminhoBanner = caminho
            Exit Function
        End If
    End If
    pasta = ThisWorkbook.Path
    If pasta = "" Then Exit Function
    If LCase$(Left$(pasta, 4)) = "http" Then Exit Function
    caminho = pasta & Application.PathSeparator & CB_ARQ_BANNER
    If Dir(caminho) <> "" Then CB_CaminhoBanner = caminho
End Function

Private Function CB_Tabela(ByVal titulo As String, ByVal itens As Object) As String
    Dim h As String, i As Long, it As Variant
    If itens.Count = 0 Then Exit Function

    h = "<p><b>" & titulo & " (" & itens.Count & ")</b></p>" & _
        "<table border=""1"" cellspacing=""0"" cellpadding=""4"" " & _
        "style=""border-collapse:collapse;font-family:Calibri,Arial;font-size:10pt"">" & _
        "<tr style=""background:#1F4E79;color:#FFFFFF"">" & _
        "<th>Chamado</th><th>C&oacute;d. Pe&ccedil;a</th><th>Descri&ccedil;&atilde;o</th>" & _
        "<th>Qtd</th><th>M&ecirc;s ref.</th><th>Prazo</th><th>Situa&ccedil;&atilde;o</th>"
    If CB_MOSTRAR_VALORES Then h = h & "<th>Valor</th>"
    h = h & "</tr>"

    For i = 1 To itens.Count
        it = itens(i)
        h = h & "<tr>" & _
            "<td>" & CB_Html(CStr(it(0))) & "</td>" & _
            "<td>" & CB_Html(CStr(it(1))) & "</td>" & _
            "<td>" & CB_Html(CStr(it(2))) & "</td>" & _
            "<td align=""center"">" & it(3) & "</td>" & _
            "<td align=""center"">" & Format$(it(4), "mm/yyyy") & "</td>" & _
            "<td align=""center"">" & CB_DataTxt(it(5)) & "</td>" & _
            "<td align=""center"">" & CB_Html(CStr(it(6))) & "</td>"
        If CB_MOSTRAR_VALORES Then _
            h = h & "<td align=""right"">" & Format$(it(7), "#,##0.00") & "</td>"
        h = h & "</tr>"
    Next i
    CB_Tabela = h & "</table><br>"
End Function

'=====================================================================
' 2) PLANILHA DE DEBITO (fim do mes)
'=====================================================================
Private Sub CB_Debito(ByVal enviar As Boolean)
    Dim dados As Object, k As Variant, info As Object, it As Variant
    Dim wbNovo As Workbook, wsNovo As Worksheet
    Dim lin As Long, i As Long, nItens As Long
    Dim total As Double, arq As String, pasta As String
    Dim ol As Object, mail As Object, msg As String, mesesArq As String
    Dim telaAnt As Boolean
    Dim regiaoNomeDeb As String, emailsDeb As String

    Set dados = CB_Pendencias()

    telaAnt = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbNovo = Workbooks.Add(xlWBATWorksheet)
    Set wsNovo = wbNovo.Worksheets(1)
    wsNovo.Name = "Debito"

    wsNovo.Range("A1:O1").Value = Array("Cod STA", "STA", "UF", "Coordenador", "Regiao", _
        "Chamado", "Cod Peca", "Descricao", "QTD", "Valor Unit", "Valor Total", _
        "Mes referencia", "Prazo", "Dias em atraso", "Mes devido")
    lin = 2

    For Each k In dados.Keys
        Set info = dados(k)
        CB_ObterRegiaoEConsultores CStr(info("uf")), CStr(info("regiao")), regiaoNomeDeb, emailsDeb
        
        For i = 1 To info("vencidas").Count
            it = info("vencidas")(i)
            wsNovo.Cells(lin, 1).Value = info("codigo")
            wsNovo.Cells(lin, 2).Value = info("nome")
            wsNovo.Cells(lin, 3).Value = info("uf")
            wsNovo.Cells(lin, 4).Value = info("coord")
            wsNovo.Cells(lin, 5).Value = regiaoNomeDeb
            wsNovo.Cells(lin, 6).Value = it(0)
            wsNovo.Cells(lin, 7).Value = it(1)
            wsNovo.Cells(lin, 8).Value = it(2)
            wsNovo.Cells(lin, 9).Value = it(3)
            wsNovo.Cells(lin, 10).Value = it(7)
            wsNovo.Cells(lin, 11).Value = it(8)
            wsNovo.Cells(lin, 12).Value = Format$(it(4), "mm/yyyy")
            wsNovo.Cells(lin, 13).Value = it(5)
            wsNovo.Cells(lin, 14).Value = CLng(CB_Hoje()) - CLng(it(5))
            wsNovo.Cells(lin, 15).Value = Format$(it(4), "yyyy-mm")
            If InStr(mesesArq, Format$(it(4), "mm/yyyy")) = 0 Then _
                mesesArq = mesesArq & Format$(it(4), "mm/yyyy") & "  "
            total = total + it(8)
            lin = lin + 1
            nItens = nItens + 1
        Next i
    Next k

    With wsNovo.Range("A1:O1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    If lin > 2 Then
        wsNovo.Range("L2:M" & lin - 1).NumberFormat = "dd/mm/yyyy"
        wsNovo.Range("J2:K" & lin - 1).NumberFormat = "#,##0.00"
        wsNovo.Range("A1:O" & lin - 1).AutoFilter
        wsNovo.Cells(lin + 1, 10).Value = "TOTAL"
        wsNovo.Cells(lin + 1, 11).Value = total
        wsNovo.Cells(lin + 1, 10).Resize(1, 2).Font.Bold = True
        wsNovo.Cells(lin + 1, 11).NumberFormat = "#,##0.00"
    End If
    wsNovo.Columns("A:O").AutoFit
    If lin > 2 Then CB_ResumoDebito wbNovo, wsNovo, lin - 1

    pasta = CB_PastaDebito()
    arq = pasta & "\Debito_RPTG_" & Format$(CB_Hoje(), "yyyy_mm") & ".xlsx"
    wbNovo.SaveAs Filename:=arq, FileFormat:=51
    wbNovo.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' e-mail para o Financeiro
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If Not ol Is Nothing And CB_EmailFinanceiro() <> "" Then
        Set mail = ol.CreateItem(0)
        mail.To = CB_EmailFinanceiro()
        mail.Subject = "RPTG - Pecas para debito - " & Format$(CB_Hoje(), "mm/yyyy")
        mail.HTMLBody = "<html><body style=""font-family:Calibri,Arial;font-size:11pt"">" & _
            "<p>Boa tarde, Financeiro.</p>" & _
            "<p>Segue a planilha das pe&ccedil;as de RPTG que passaram do prazo de " & _
            "devolu&ccedil;&atilde;o (dia " & CB_DIA_LIMITE & ") e devem ser <b>debitadas</b>:</p>" & _
            "<ul><li>Meses de refer&ecirc;ncia: <b>" & mesesArq & "</b></li>" & _
            "<li>Pe&ccedil;as: <b>" & nItens & "</b></li>" & _
            "<li>Valor total: <b>R$ " & Format$(total, "#,##0.00") & "</b></li></ul>" & _
            "<p>O arquivo tem a aba <b>Resumo</b> (tabela din&acirc;mica por m&ecirc;s devido " & _
            "e por STA) e a aba <b>Debito</b> com o detalhe pe&ccedil;a por pe&ccedil;a.</p>" & _
            "<p>STAs MEI e da regi&atilde;o Norte n&atilde;o entram na cobran&ccedil;a.</p>" & _
            CB_Assinatura() & "</body></html>"
        CB_AnexarBanner mail
        mail.Attachments.Add arq
        If enviar Then mail.Send Else mail.Display
    End If

    msg = "Planilha de debito: " & nItens & " peca(s), R$ " & Format$(total, "#,##0.00") & _
          vbCrLf & "Meses no arquivo: " & mesesArq & vbCrLf & arq
    CB_Log msg
    Application.ScreenUpdating = telaAnt
    If Not enviar Then MsgBox msg, vbInformation, "Planilha de debito"
End Sub

'=====================================================================
' LEITURA DA BASE
'=====================================================================
Private Function CB_Pendencias() As Object
    Dim ws As Worksheet, lo As ListObject
    Dim dic As Object, dSta As Object, info As Object
    Dim v As Variant, i As Long
    Dim cCham As Long, cSta As Long, cNome As Long, cPeca As Long, cDesc As Long
    Dim cQtd As Long, cEnc As Long, cStatus As Long, cVal As Long, cTot As Long
    Dim cMEI As Long, cUF As Long, cCoord As Long, cMes As Long, cAno As Long
    Dim cod As String, st As String, prazo As Date, item As Variant
    Dim mesRef As Long, anoRef As Long

    Set dic = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CB_ABA_BASE)
    Set lo = ws.ListObjects(CB_NOME_TAB)
    If lo.ListRows.Count = 0 Then
        Set CB_Pendencias = dic
        Exit Function
    End If

    cCham = CB_Col(lo, "Chamado"): cSta = CB_Col(lo, "Codigo_STA")
    cNome = CB_Col(lo, "STA"): cPeca = CB_Col(lo, "Codigo_Peca")
    cDesc = CB_Col(lo, "Descricao"): cQtd = CB_Col(lo, "QTD")
    cEnc = CB_Col(lo, "Dta_Encerramento"): cStatus = CB_Col(lo, "Status")
    cVal = CB_Col(lo, "Valor_Debito_Unit"): cTot = CB_Col(lo, "Valor_Total")
    cMEI = CB_Col(lo, "MEI"): cUF = CB_Col(lo, "UF"): cCoord = CB_Col(lo, "Coordenador")
    cMes = CB_Col(lo, "Mes"): cAno = CB_Col(lo, "Ano")

    Set dSta = CB_DicSTAs()
    v = lo.DataBodyRange.Value

    For i = 1 To UBound(v, 1)
        cod = CB_SoNum(v(i, cSta))
        st = CB_Norm(CB_Txt(v(i, cStatus)))
        mesRef = 0: anoRef = 0
        If cMes > 0 Then mesRef = CLng(CB_Num(v(i, cMes)))
        If cAno > 0 Then anoRef = CLng(CB_Num(v(i, cAno)))
        If (mesRef < 1 Or mesRef > 12 Or anoRef < 1900) And IsDate(v(i, cEnc)) Then
            mesRef = Month(CDate(v(i, cEnc)))
            anoRef = Year(CDate(v(i, cEnc)))
        End If
        If cod <> "" And CB_Cobravel(st) And mesRef >= 1 And mesRef <= 12 And anoRef >= 1900 Then
          If CB_MesAceito(DateSerial(anoRef, mesRef, 1)) Then
            If Not CB_Excluir(cod, dSta, CB_Txt(v(i, cMEI)), CB_Txt(v(i, cCoord))) Then
                prazo = CB_PrazoMes(anoRef, mesRef)
                If Not dic.Exists(cod) Then
                    Set info = CreateObject("Scripting.Dictionary")
                    info("codigo") = cod
                    info("nome") = CB_Info(dSta, cod, "nome", CB_Txt(v(i, cNome)))
                    info("email") = CB_Info(dSta, cod, "email", "")
                    info("uf") = CB_Info(dSta, cod, "uf", CB_Txt(v(i, cUF)))
                    info("coord") = CB_Info(dSta, cod, "coord", CB_Txt(v(i, cCoord)))
                    info("regiao") = CB_Info(dSta, cod, "regiao", "")
                    Set info("dentro") = New Collection
                    Set info("vencidas") = New Collection
                    dic.Add cod, info
                End If
                item = Array(CB_Txt(v(i, cCham)), CB_Txt(v(i, cPeca)), CB_Txt(v(i, cDesc)), _
                             CB_Num(v(i, cQtd)), DateSerial(anoRef, mesRef, 1), prazo, _
                             CB_Txt(v(i, cStatus)), CB_Num(v(i, cVal)), CB_Num(v(i, cTot)))
                If CB_Hoje() > prazo Or CB_MesEscolhido() Then
                    dic(cod)("vencidas").Add item
                Else
                    dic(cod)("dentro").Add item
                End If
            End If
          End If
        End If
    Next i

    Set CB_Pendencias = dic
End Function

Private Function CB_PrazoMes(ByVal ano As Long, ByVal mes As Long) As Date
    CB_PrazoMes = DateSerial(ano, mes + CB_MESES, CB_DIA_LIMITE)
End Function

Private Function CB_Prazo(ByVal dtaEnc As Date) As Date
    CB_Prazo = CB_PrazoMes(Year(dtaEnc), Month(dtaEnc))
End Function

Private Function CB_Cobravel(ByVal statusNorm As String) As Boolean
    CB_Cobravel = (statusNorm = "aguardandodevolucao")
End Function

Private Function CB_Excluir(ByVal cod As String, ByVal dSta As Object, _
                            ByVal meiBase As String, ByVal coordBase As String) As Boolean
    Dim mei As String, reg As String
    mei = CB_Norm(CB_Info(dSta, cod, "mei", meiBase))
    reg = CB_Norm(CB_Info(dSta, cod, "regiao", ""))
    If reg = "" Then reg = CB_Norm(CB_Info(dSta, cod, "coord", coordBase))
    CB_Excluir = (mei = "sim") Or (reg = "norte")
End Function

Private Function CB_DicSTAs() As Object
    Dim ws As Worksheet, d As Object, ult As Long, i As Long, cod As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CB_ABA_STAS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set CB_DicSTAs = d
        Exit Function
    End If
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = CB_SoNum(ws.Cells(i, 6).Value)
        If cod <> "" And Not d.Exists(cod) Then
            d.Add cod, Array(CB_Txt(ws.Cells(i, 3).Value), _
                             CB_Txt(ws.Cells(i, 7).Value), _
                             CB_Txt(ws.Cells(i, 1).Value), _
                             CB_Txt(ws.Cells(i, 5).Value), _
                             CB_Txt(ws.Cells(i, 4).Value), _
                             CB_Txt(ws.Cells(i, 10).Value))
        End If
    Next i
    Set CB_DicSTAs = d
End Function

Private Function CB_Info(ByVal dSta As Object, ByVal cod As String, _
                         ByVal campo As String, ByVal padrao As String) As String
    Dim a As Variant, idx As Long
    If Not dSta.Exists(cod) Then
        CB_Info = padrao
        Exit Function
    End If
    a = dSta(cod)
    Select Case campo
        Case "nome": idx = 0
        Case "email": idx = 1
        Case "uf": idx = 2
        Case "coord": idx = 3
        Case "regiao": idx = 4
        Case "mei": idx = 5
        Case Else: idx = -1
    End Select
    If idx < 0 Then
        CB_Info = padrao
    ElseIf Trim$(CStr(a(idx))) = "" Then
        CB_Info = padrao
    Else
        CB_Info = Trim$(CStr(a(idx)))
    End If
End Function

'=====================================================================
' MAPEAMENTO DE REGIOES E CONSULTORES EM COPIA (CC)
'=====================================================================
Private Sub CB_ObterRegiaoEConsultores(ByVal uf As String, ByVal regiaoBase As String, ByRef regiaoNome As String, ByRef emailsCC As String)
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
            If Trim$(regiaoBase) <> "" Then
                regiaoNome = UCase$(Trim$(regiaoBase))
            Else
                regiaoNome = "GERAL"
            End If
            emailsPadrao = ""
    End Select

    ' Tenta buscar se o usuario personalizou o e-mail na tabela da aba Parametros
    emailsCC = ""
    On Error Resume Next
    Set wsP = CB_Param()
    On Error GoTo 0
    
    If Not wsP Is Nothing Then
        ultP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
        If ultP < 25 Then ultP = 25
        For i = 1 To ultP
            regTab = UCase$(Trim$(CStr(wsP.Cells(i, 1).Value)))
            If regTab = "" Then regTab = UCase$(Trim$(CStr(wsP.Cells(i, 15).Value)))
            
            If regTab <> "" Then
                If regTab = UCase$(regiaoNome) Or InStr(1, regTab, UCase$(regiaoNome)) > 0 Or _
                   (regiaoNome = "SUDESTE RJ/ES" And (InStr(1, regTab, "RJ") > 0 Or InStr(1, regTab, "ES") > 0)) Then
                    If InStr(1, CStr(wsP.Cells(i, 3).Value), "@") > 0 Then
                        emailsCC = Trim$(CStr(wsP.Cells(i, 3).Value))
                    ElseIf InStr(1, CStr(wsP.Cells(i, 16).Value), "@") > 0 Then
                        emailsCC = Trim$(CStr(wsP.Cells(i, 16).Value))
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

' Cc do e-mail de cobranca: consultores da regiao/UF + Financeiro (Parametros P20)
Private Function CB_Copias(ByVal emailsConsultores As String) As String
    Dim cc As String, financeiro As String
    cc = emailsConsultores
    financeiro = CB_EmailFinanceiro()
    If financeiro <> "" Then
        If InStr(1, ";" & cc & ";", ";" & financeiro & ";", vbTextCompare) = 0 And InStr(1, cc, financeiro, vbTextCompare) = 0 Then
            If cc <> "" Then cc = cc & "; "
            cc = cc & financeiro
        End If
    End If
    CB_Copias = cc
End Function

'=====================================================================
' UTILITARIOS
'=====================================================================
Private Function CB_Param() As Worksheet
    On Error Resume Next
    Set CB_Param = ThisWorkbook.Worksheets(CB_ABA_PARAM)
End Function

Private Function CB_EmailFinanceiro() As String
    Dim ws As Worksheet
    Set ws = CB_Param()
    If ws Is Nothing Then Exit Function
    CB_EmailFinanceiro = Trim$(CStr(ws.Range(CB_CEL_FINANCEIRO).Value & ""))
End Function

' tabela dinamica no arquivo do Financeiro: mes devido / STA x pecas e valor
Private Sub CB_ResumoDebito(wbNovo As Workbook, wsDados As Worksheet, ByVal ultLin As Long)
    Dim wsRes As Worksheet, pc As PivotCache, pt As PivotTable

    On Error Resume Next
    Set wsRes = wbNovo.Worksheets.Add(Before:=wsDados)
    If wsRes Is Nothing Then Exit Sub
    wsRes.Name = "Resumo"

    Set pc = wbNovo.PivotCaches.Create(xlDatabase, _
             wsDados.Range(wsDados.Cells(1, 1), wsDados.Cells(ultLin, 15)))
    Set pt = pc.CreatePivotTable(wsRes.Range("A4"), "ptDebito")
    If pt Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    With pt
        .AddFields RowFields:=Array("Mes devido", "Cod STA", "STA")
        .PivotFields("Regiao").Orientation = xlPageField
        With .PivotFields("QTD")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "Pecas "
            .NumberFormat = "#,##0"
        End With
        With .PivotFields("Valor Total")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "Valor "
            .NumberFormat = "#,##0.00"
        End With
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium2"
    End With

    wsRes.Range("A1").Value = "RPTG - pecas para debito - " & Format$(CB_Hoje(), "mm/yyyy")
    wsRes.Range("A2").Value = "Mes devido = mes de encerramento da peca (prazo vencido no dia " & _
                              CB_DIA_LIMITE & " do " & CB_MESES & "o mes seguinte). " & _
                              "Fora MEI e regiao Norte."
    wsRes.Range("A1").Font.Bold = True
    wsRes.Range("A1").Font.Size = 14
    wsRes.Columns("A:F").AutoFit
    On Error GoTo 0
End Sub

Private Function CB_PastaDebito() As String
    Dim ws As Worksheet, p As String
    Set ws = CB_Param()
    If Not ws Is Nothing Then p = Trim$(CStr(ws.Range(CB_CEL_PASTA).Value & ""))
    If p = "" Then p = ThisWorkbook.Path
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    CB_PastaDebito = p
End Function

Private Function CB_Col(ByVal lo As ListObject, ByVal nome As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If CB_Norm(lo.ListColumns(i).Name) = CB_Norm(nome) Then
            CB_Col = i
            Exit Function
        End If
    Next i
End Function

Private Function CB_SoNum(ByVal v As Variant) As String
    Dim i As Long, ch As String, r As String, T As String
    T = CB_Txt(v)
    For i = 1 To Len(T)
        ch = Mid$(T, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    If r <> "" Then r = CStr(CDbl(r))
    CB_SoNum = r
End Function

Private Function CB_Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CB_Txt = Trim$(CStr(v))
End Function

Private Function CB_Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then CB_Num = CDbl(v)
End Function

Private Function CB_DataTxt(ByVal v As Variant) As String
    If IsDate(v) Then CB_DataTxt = Format$(CDate(v), "dd/mm/yyyy")
End Function

Private Function CB_Html(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    CB_Html = s
End Function

Private Function CB_Norm(ByVal s As String) As String
    Dim i As Long, cod As Long, ch As String, r As String
    s = Trim$(s)
    For i = 1 To Len(s)
        cod = AscW(Mid$(s, i, 1))
        If cod < 0 Then cod = cod + 65536
        Select Case cod
            Case 192 To 197, 224 To 229: ch = "a"
            Case 199, 231:               ch = "c"
            Case 200 To 203, 232 To 235: ch = "e"
            Case 204 To 207, 236 To 239: ch = "i"
            Case 209, 241:               ch = "n"
            Case 210 To 214, 242 To 246: ch = "o"
            Case 217 To 220, 249 To 252: ch = "u"
            Case 221, 253, 255:          ch = "y"
            Case Else:                   ch = LCase$(Mid$(s, i, 1))
        End Select
        If (ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Then r = r & ch
    Next i
    CB_Norm = r
End Function

Private Sub CB_Log(ByVal msg As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open ThisWorkbook.Path & "\" & CB_ARQ_LOG For Append As #f
    Print #f, Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & Replace(msg, vbCrLf, " | ")
    Close #f
End Sub

Public Sub Destravar_Tela_Cobranca()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Tela liberada.", vbInformation
End Sub