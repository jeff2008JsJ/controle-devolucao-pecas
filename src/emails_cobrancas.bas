Attribute VB_Name = "emails_cobrancas"
Option Explicit

' ============================================================================
'  Envio dos e-mails de cobranca de devolucao RPTG para os STAs pendentes.
'  Aba Emails_Devolucao: A = e-mail, B = assunto, C = corpo, I = log do envio.
'  O corpo e enviado em HTML, com a assinatura e o banner da Empresa.
' ============================================================================

Private Const ABA_EMAILS As String = "Emails_Devolucao"   ' aba propria (nao mexe na SendBulkEmails da coleta)
Private Const ABA_BASE As String = "Base_RPTG"
Private Const COL_LOG As Long = 9              ' coluna I: registro do envio
Private Const COL_DETALHE As Long = 11         ' coluna K: meses em aberto
Private Const COL_REGRA As Long = 12           ' coluna L: regra de cobranca
Private Const COL_DEBITO As Long = 13          ' coluna M: valor a debitar
Private Const MESES_LIMITE As Long = 3         ' acima disso, debita do STA
Private Const UFS_NORTE As String = "AC,AP,AM,PA,RO,RR,TO"   ' isentas de debito
Private Const PARA_FIXO As String = "financeiro@exemplo.com.br"   ' sempre no "Para"
Private Const COPIA_PARA As String = "retorno.pecas@exemplo.com.br" ' sempre em copia
Private Const PULAR_JA_ENVIADOS As Boolean = True
Private Const PAUSA_SEGUNDOS As Double = 1     ' pausa entre envios
Private Const MAX_REVISAR As Long = 5          ' quantos abrir na revisao (0 = todos)
Private Const ARQ_BANNER As String = "Assinatura_Empresa.png"
Private Const CAMINHO_BANNER As String = ""       ' opcional: caminho completo do .png
Private Const ABA_PARAM As String = "Parametros"  ' onde o caminho fica guardado
Private Const CEL_BANNER As String = "P6"
Private Const LARGURA_BANNER As Long = 500     ' largura do banner no e-mail (px)
Private Const DIAS_CICLO As Long = 15          ' intervalo do envio quinzenal
Private Const DIAS_CICLO_SEMANAL As Long = 7   ' intervalo do envio semanal
Private Const ABA_STAS As String = "STAs"      ' cadastro dos STAs (MEI? / Regiao)
Private Const ARQ_LOG As String = "Log_Envio_RPTG.txt"
Private Const CEL_FINANCEIRO As String = "P20"      ' e-mail do Financeiro (Cc de todos)
Private Const PAR_COL_REG As Long = 15         ' Parametros coluna O: regiao (ou regiao + UF)
Private Const PAR_COL_CC As Long = 16          ' Parametros coluna P: consultores em Cc
Private Const CB_DIA_LIMITE As Long = 15       ' prazo: dia 15 ...
Private Const CB_MESES As Long = 3             ' ... do 3o mes depois do mes de referencia
Private Const SO_MESES_NO_PRAZO As Boolean = True ' cobra so os meses ainda dentro do prazo

' quando True, nenhuma janela e mostrada (usado no envio agendado)
Private silencioso As Boolean
' STAs isentas de cobranca (MEI e regiao Norte)
Private dicIsentas As Object
' regiao|UF de cada STA, para o Cc dos consultores
Private dicRegUf As Object

' --- usados apenas pela macro Cobrar_Meses_Escolhidos ---
Private mesesEscolhidos As String
Private prazoManual As Date
Private ignorarEnviados As Boolean

Private Function AbaEmails() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_EMAILS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ABA_EMAILS
    End If
    Set AbaEmails = ws
End Function

' ---------------------------------------------------------------------------
'  Cobrança de meses escolhidos manualmente
' ---------------------------------------------------------------------------
Public Sub Cobrar_Meses_Escolhidos()
    Dim r As String, partes As Variant, i As Long
    Dim mes As Long, ano As Long, lista As String, mostrar As String
    Dim enviar As VbMsgBoxResult

    r = InputBox(T("Quais meses cobrar? Separe por ; - ex.: 04/2026;05/2026") & vbCrLf & _
                 T("(pode digitar s[o'] o m[e^]s: 4;5 usa o ano atual)"), _
                 T("Cobran[c]a de meses escolhidos"))
    If Trim$(r) = "" Then Exit Sub

    partes = Split(Replace(Replace(r, ",", ";"), " ", ""), ";")
    For i = LBound(partes) To UBound(partes)
        If Trim$(CStr(partes(i))) <> "" Then
            If InStr(CStr(partes(i)), "/") > 0 Then
                mes = Val(Split(CStr(partes(i)), "/")(0))
                ano = Val(Split(CStr(partes(i)), "/")(1))
            Else
                mes = Val(CStr(partes(i)))
                ano = Year(Date)
            End If
            If mes >= 1 And mes <= 12 And ano >= 1900 Then
                lista = lista & Format$(ano, "0000") & "-" & Format$(mes, "00") & ";"
                mostrar = mostrar & Format$(mes, "00") & "/" & Format$(ano, "0000") & "  "
            End If
        End If
    Next i

    If lista = "" Then
        MsgBox T("N[a~]o entendi os meses. Use o formato 04/2026;05/2026."), vbExclamation
        Exit Sub
    End If

    r = InputBox(T("Devolver as pe[c]as at[e'] qual data? (dd/mm/aaaa)"), _
                 T("Cobran[c]a de meses escolhidos"), Format$(Date + 30, "dd/mm/yyyy"))
    If Trim$(r) = "" Then Exit Sub
    If Not IsDate(r) Then
        MsgBox T("Data inv[a']lida. Use dd/mm/aaaa, ex.: 31/07/2026."), vbExclamation
        Exit Sub
    End If

    enviar = MsgBox(T("Meses: ") & mostrar & vbCrLf & _
                    T("Prazo para devolver: ") & Format$(CDate(r), "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                    T("Enviar os e-mails agora?") & vbCrLf & _
                    T("Sim = envia pelo Outlook   /   N[a~]o = abre para revisar"), _
                    vbYesNoCancel + vbQuestion, T("Cobran[c]a de meses escolhidos"))
    If enviar = vbCancel Then Exit Sub

    mesesEscolhidos = ";" & lista
    prazoManual = CDate(r)
    ignorarEnviados = True
    silencioso = True

    On Error GoTo limpar
    Atualizar_Emails_STA
    silencioso = False
    ProcessarEmails (enviar = vbYes), ""

limpar:
    mesesEscolhidos = ""
    prazoManual = 0
    ignorarEnviados = False
    If Err.Number <> 0 Then
        MsgBox "Erro " & Err.Number & ": " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error Resume Next
    silencioso = True
    Atualizar_Emails_STA
    silencioso = False
    On Error GoTo 0
End Sub

Public Sub Revisar_Emails_STA()
    ProcessarEmails False, ""
End Sub

Public Sub Liberar_Envios()
    Dim ws As Worksheet, i As Long, ult As Long, n As Long

    Set ws = AbaEmails()
    ult = UltimaLinha(ws)
    If MsgBox(T("Limpar a coluna I (") & T("[u']ltimo envio") & ") de todos os STAs," & vbCrLf & _
              T("para eles voltarem a ser cobrados?"), vbYesNo + vbQuestion, _
              T("Devolu[c][a~]o RPTG")) <> vbYes Then Exit Sub

    For i = 2 To ult
        If Trim(CStr(ws.Cells(i, COL_LOG).Value)) <> "" Then
            ws.Cells(i, COL_LOG).ClearContents
            n = n + 1
        End If
    Next i
    MsgBox n & T(" STA(s) liberados para nova cobran[c]a."), vbInformation
End Sub

Public Sub Enviar_Emails_STA()
    Dim n As Long
    n = ContarPendentes()
    If n = 0 Then
        MsgBox "Nenhum STA pendente a enviar na aba " & ABA_EMAILS & ".", vbInformation
        Exit Sub
    End If
    If MsgBox("Enviar " & n & " e-mail(s) agora?", vbYesNo + vbQuestion, T("Devolu[c][a~]o RPTG")) <> vbYes Then Exit Sub
    ProcessarEmails True, ""
End Sub

Public Sub Envio_Semanal_Devolucao()
    EnvioAutomatico DIAS_CICLO_SEMANAL
End Sub

Public Sub Envio_Quinzenal_Automatico()
    EnvioAutomatico DIAS_CICLO
End Sub

Private Sub EnvioAutomatico(ByVal dias As Long)
    Dim liberadas As Long
    silencioso = True
    On Error GoTo falha

    Application.DisplayAlerts = False
    Atualizar_Meses_Devidos
    liberadas = LiberarEnviosAntigos(dias)
    Gravar "inicio do envio automatico; STAs liberadas: " & liberadas
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

Private Function LiberarEnviosAntigos(ByVal dias As Long) As Long
    Dim ws As Worksheet, i As Long, ult As Long, n As Long
    Dim v As Variant

    Set ws = AbaEmails()
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

Public Sub Enviar_Email_Teste()
    Dim destino As String
    destino = InputBox(T("Enviar o primeiro e-mail para qual endere[c]o de teste?"), "Teste de envio")
    If Trim(destino) = "" Then Exit Sub
    ProcessarEmails True, destino
End Sub

Private Function UltimaLinha(ws As Worksheet) As Long
    Dim a As Long, d As Long
    a = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    d = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    UltimaLinha = IIf(a > d, a, d)
End Function

Private Function LinhaValida(ws As Worksheet, ByVal i As Long) As Boolean
    Dim cod As String
    If InStr(1, CStr(ws.Cells(i, 1).Value), "@") = 0 Then Exit Function
    If PULAR_JA_ENVIADOS And Not ignorarEnviados Then
        If Trim(CStr(ws.Cells(i, COL_LOG).Value)) <> "" Then Exit Function
    End If
    cod = SoNumero(CStr(ws.Cells(i, 4).Value))
    If StaIsenta(cod) Then Exit Function
    If Val(CStr(ws.Cells(i, 6).Value)) <= 0 Then Exit Function
    LinhaValida = True
End Function

Private Function Cobravel(ByVal status As String) As Boolean
    Cobravel = (Normalizar(status) = "aguardandodevolucao")
End Function

Private Function StaIsenta(ByVal cod As String) As Boolean
    If dicIsentas Is Nothing Then CarregarIsentas
    If dicIsentas Is Nothing Then Exit Function
    StaIsenta = dicIsentas.Exists(SoNumero(cod))
End Function

Private Sub CarregarIsentas()
    Dim ws As Worksheet, dados As Variant, i As Long, ult As Long
    Dim cCod As Long, cMEI As Long, cReg As Long, cUF As Long, cod As String
    Dim reg As String, uf As String

    Set dicIsentas = CreateObject("Scripting.Dictionary")
    Set dicRegUf = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_STAS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    cCod = ColunaN(ws, "codsta")
    cMEI = ColunaN(ws, "mei")
    cReg = ColunaN(ws, "regiao")
    cUF = ColunaN(ws, "uf")
    If cCod = 0 Then Exit Sub

    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Sub
    dados = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    For i = 1 To UBound(dados, 1)
        cod = SoNumero(Txt(dados(i, cCod)))
        If cod <> "" Then
            reg = ""
            uf = ""
            If cReg > 0 Then reg = Txt(dados(i, cReg))
            If cUF > 0 Then uf = Txt(dados(i, cUF))
            dicRegUf(cod) = reg & "|" & uf
            If cMEI > 0 Then
                If Left$(Normalizar(Txt(dados(i, cMEI))), 1) = "s" Then dicIsentas(cod) = True
            End If
            If cReg > 0 Then
                If Normalizar(Txt(dados(i, cReg))) = "norte" Then dicIsentas(cod) = True
            End If
            If cUF > 0 Then
                If InStr(1, "," & UFS_NORTE & ",", "," & UCase$(Txt(dados(i, cUF))) & ",") > 0 Then _
                    dicIsentas(cod) = True
            End If
        End If
    Next i
End Sub

Private Function ColunaN(ws As Worksheet, ByVal tituloNorm As String) As Long
    Dim j As Long
    For j = 1 To ws.UsedRange.Columns.Count
        If Normalizar(CStr(ws.Cells(1, j).Value)) = tituloNorm Then
            ColunaN = j
            Exit Function
        End If
    Next j
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function

Private Function SoNumero(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    Do While Len(r) > 1 And Left$(r, 1) = "0"
        r = Mid$(r, 2)
    Loop
    SoNumero = r
End Function

Private Function Normalizar(ByVal s As String) As String
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
    Normalizar = r
End Function

Private Function ContarPendentes() As Long
    Dim ws As Worksheet, i As Long, ult As Long, n As Long
    Set ws = AbaEmails()
    ult = UltimaLinha(ws)
    For i = 2 To ult
        If LinhaValida(ws, i) Then n = n + 1
    Next i
    ContarPendentes = n
End Function

Public Sub Atualizar_Aging_Devolucao()
    Dim ws As Worksheet
    Dim cGrp As Long, cDta As Long, cDias As Long, cFaixa As Long, cCham As Long
    Dim dados As Variant, dias As Variant, faixa As Variant
    Dim i As Long, ult As Long, n As Long, d As Double, hoje As Double
    Dim grp As String, calcAntes As XlCalculation

    Set ws = ThisWorkbook.Worksheets(ABA_BASE)
    cGrp = Coluna(ws, "Status")
    cDta = Coluna(ws, "Dta_Encerramento")
    cDias = Coluna(ws, "Dias_Em_Aberto")
    cFaixa = Coluna(ws, "Faixa_Aging")
    cCham = Coluna(ws, "Chamado")
    If cGrp = 0 Or cDta = 0 Or cDias = 0 Or cFaixa = 0 Or cCham = 0 Then
        If silencioso Then
            Gravar "nao encontrei as colunas de aging na aba " & ABA_BASE
        Else
            MsgBox T("N[a~]o encontrei as colunas de aging na aba ") & ABA_BASE & ".", vbCritical
        End If
        Exit Sub
    End If

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
        grp = Txt(dados(i, cGrp))
        If Cobravel(grp) Then
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

    If Not silencioso Then
        MsgBox "Aging atualizado em " & Format$(n, "#,##0") & " linha(s).", vbInformation
    End If
End Sub

Public Sub Atualizar_Meses_Devidos()
    Dim wsB As Worksheet, wsE As Worksheet
    Dim dic As Object, chave As String, cod As String
    Dim dados As Variant, i As Long, ult As Long
    Dim cCod As Long, cQtd As Long, cVal As Long, cMes As Long, cAno As Long, cGrp As Long
    Dim mes As Long, ano As Long, periodos As Object, itens As Variant
    Dim texto As String, k As Variant
    Dim cMEI As Long, cUF As Long
    Dim dicMei As Object, dicNorte As Object
    Dim refMes As Long, debito As Double, regra As String, motivos As String

    Atualizar_Aging_Devolucao
    Atualizar_Emails_STA

    Set wsB = ThisWorkbook.Worksheets(ABA_BASE)
    Set wsE = AbaEmails()

    cCod = Coluna(wsB, "Codigo_STA")
    cQtd = Coluna(wsB, "QTD")
    cVal = Coluna(wsB, "Valor_Total")
    cMes = Coluna(wsB, "Mes")
    cAno = Coluna(wsB, "Ano")
    cGrp = Coluna(wsB, "Status")
    cMEI = Coluna(wsB, "MEI")
    cUF = Coluna(wsB, "UF")
    If cCod = 0 Or cQtd = 0 Or cVal = 0 Or cMes = 0 Or cAno = 0 Or cGrp = 0 _
       Or cMEI = 0 Or cUF = 0 Then
        If silencioso Then
            Gravar "nao encontrei as colunas necessarias na aba " & ABA_BASE
        Else
            MsgBox T("N[a~]o encontrei as colunas necess[a']rias na aba ") & ABA_BASE & ".", vbCritical
        End If
        Exit Sub
    End If

    ult = wsB.Cells(wsB.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Set dic = CreateObject("Scripting.Dictionary")
    Set dicMei = CreateObject("Scripting.Dictionary")
    Set dicNorte = CreateObject("Scripting.Dictionary")
    refMes = Year(Date) * 12 + Month(Date)
    dados = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, wsB.UsedRange.Columns.Count)).Value

    For i = 1 To UBound(dados, 1)
        If Cobravel(Txt(dados(i, cGrp))) Then
            cod = SoNumero(Txt(dados(i, cCod)))
            If cod <> "" And Not StaIsenta(cod) Then
                ano = Num(dados(i, cAno))
                mes = Num(dados(i, cMes))
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
                itens = Array(itens(0) + Num(dados(i, cQtd)), itens(1) + Num(dados(i, cVal)))
                periodos(chave) = itens
                If UCase$(Left$(Txt(dados(i, cMEI)), 1)) = "S" Then dicMei(cod) = True
                If InStr(1, "," & UFS_NORTE & ",", "," & UCase$(Txt(dados(i, cUF))) & ",") > 0 Then dicNorte(cod) = True
            End If
        End If
    Next i

    ult = UltimaLinha(wsE)
    wsE.Cells(1, COL_DETALHE).Value = T("Meses em aberto (pe[c]as e valor)")
    wsE.Cells(1, COL_REGRA).Value = T("Regra de cobran[c]a")
    wsE.Cells(1, COL_DEBITO).Value = "Valor em aberto (dentro do prazo)"
    If ult >= 2 Then wsE.Range(wsE.Cells(2, COL_DETALHE), wsE.Cells(ult, COL_DEBITO)).ClearContents

    For i = 2 To ult
        cod = SoNumero(CStr(wsE.Cells(i, 4).Value))
        texto = ""
        debito = 0
        regra = ""
        If dic.Exists(cod) Then
            motivos = ""
            If dicMei.Exists(cod) Then motivos = "MEI"
            If dicNorte.Exists(cod) Then motivos = motivos & IIf(motivos = "", "", " e ") & T("regi[a~]o Norte")
            regra = IIf(motivos = "", T("Sujeito a d[e']bito"), T("Isento de d[e']bito - ") & motivos)

            Set periodos = dic(cod)
            For Each k In OrdenarChaves(periodos.Keys)
                itens = periodos(k)
                ano = Val(Left$(CStr(k), 4))
                mes = Val(Mid$(CStr(k), 6, 2))
                texto = texto & IIf(texto = "", "", vbLf) & _
                        Mid$(CStr(k), 6, 2) & "/" & Left$(CStr(k), 4) & ": " & _
                        Format$(itens(0), "#,##0") & T(" pe[c]a(s) - R$ ") & _
                        Format$(itens(1), "#,##0.00")
                If mes >= 1 And mes <= 12 Then
                    If Date > PrazoDevolucao(ano, mes) Then
                        If motivos = "" Then
                            texto = texto & T(" - a debitar")
                            debito = debito + itens(1)
                        Else
                            texto = texto & T(" - isento")
                        End If
                    End If
                End If
            Next k
        End If
        wsE.Cells(i, COL_DETALHE).Value = texto
        wsE.Cells(i, COL_REGRA).Value = regra
        If regra <> "" Then wsE.Cells(i, COL_DEBITO).Value = debito
    Next i
    wsE.Columns(COL_DETALHE).ColumnWidth = 46
    wsE.Columns(COL_REGRA).ColumnWidth = 30
    wsE.Columns(COL_DEBITO).ColumnWidth = 22
    wsE.Range(wsE.Cells(2, COL_DEBITO), wsE.Cells(ult, COL_DEBITO)).NumberFormat = "#,##0.00"
    Application.ScreenUpdating = True

    If silencioso Then
        Gravar "detalhamento e regra de cobranca atualizados para " & dic.Count & " STA(s)"
    Else
        MsgBox T("Detalhamento e regra de cobran[c]a atualizados para ") & dic.Count & " STA(s).", vbInformation
    End If
End Sub

Public Sub Atualizar_Emails_STA()
    Dim wsB As Worksheet, wsE As Worksheet
    Dim cEmail As Long, cCod As Long, cNome As Long, cQtd As Long, cVal As Long
    Dim cMes As Long, cAno As Long, cSts As Long, cDta As Long
    Dim dados As Variant, i As Long, ult As Long, lin As Long
    Dim dic As Object, dicEmail As Object, dicNome As Object, dicDias As Object
    Dim dicLog As Object, periodos As Object, itens As Variant
    Dim cod As String, chave As String, k As Variant, k2 As Variant
    Dim qtdTot As Double, valTot As Double, dias As Double, hoje As Double
    Dim anoRef As Long, mesRef As Long

    Set wsB = ThisWorkbook.Worksheets(ABA_BASE)
    Set wsE = AbaEmails()

    cEmail = Coluna(wsB, "Email_STA")
    cCod = Coluna(wsB, "Codigo_STA")
    cNome = Coluna(wsB, "STA")
    cQtd = Coluna(wsB, "QTD")
    cVal = Coluna(wsB, "Valor_Total")
    cMes = Coluna(wsB, "Mes")
    cAno = Coluna(wsB, "Ano")
    cSts = Coluna(wsB, "Status")
    cDta = Coluna(wsB, "Dta_Encerramento")
    If cEmail = 0 Or cCod = 0 Or cQtd = 0 Or cVal = 0 Or cMes = 0 Or cAno = 0 Or cSts = 0 Then
        If silencioso Then
            Gravar "nao encontrei as colunas necessarias para montar os e-mails"
        Else
            MsgBox T("N[a~]o encontrei as colunas necess[a']rias na aba ") & ABA_BASE & ".", vbCritical
        End If
        Exit Sub
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    Set dicEmail = CreateObject("Scripting.Dictionary")
    Set dicNome = CreateObject("Scripting.Dictionary")
    Set dicDias = CreateObject("Scripting.Dictionary")
    Set dicLog = CreateObject("Scripting.Dictionary")
    hoje = CDbl(Date)

    ult = UltimaLinha(wsE)
    For i = 2 To ult
        cod = SoNumero(CStr(wsE.Cells(i, 4).Value))
        If cod <> "" And Trim(CStr(wsE.Cells(i, COL_LOG).Value)) <> "" Then
            dicLog(cod) = wsE.Cells(i, COL_LOG).Value
        End If
    Next i

    ult = wsB.Cells(wsB.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Sub
    Application.ScreenUpdating = False
    dados = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, wsB.UsedRange.Columns.Count)).Value

    For i = 1 To UBound(dados, 1)
        If Cobravel(Txt(dados(i, cSts))) Then
            cod = SoNumero(Txt(dados(i, cCod)))
            anoRef = Num(dados(i, cAno))
            mesRef = Num(dados(i, cMes))
            If cod <> "" And Not StaIsenta(cod) And PeriodoCobravel(anoRef, mesRef) Then
                chave = Format$(anoRef, "0000") & "-" & Format$(mesRef, "00")
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
                itens = Array(itens(0) + Num(dados(i, cQtd)), itens(1) + Num(dados(i, cVal)))
                periodos(chave) = itens

                If Not dicEmail.Exists(cod) Then
                    dicEmail(cod) = Txt(dados(i, cEmail))
                    If cNome > 0 Then dicNome(cod) = Txt(dados(i, cNome))
                End If
                If cDta > 0 Then
                    If IsDate(dados(i, cDta)) Then
                        dias = hoje - CDbl(CDate(dados(i, cDta)))
                        If Not dicDias.Exists(cod) Then
                            dicDias(cod) = dias
                        ElseIf dias > CDbl(dicDias(cod)) Then
                            dicDias(cod) = dias
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ult = UltimaLinha(wsE)
    If ult >= 2 Then wsE.Range(wsE.Cells(2, 1), wsE.Cells(ult, COL_DEBITO)).ClearContents
    wsE.Cells(1, 1).Value = "EmailAddress"
    wsE.Cells(1, 2).Value = "Subject"
    wsE.Cells(1, 3).Value = "Body"
    wsE.Cells(1, 4).Value = T("C[o']d STA")
    wsE.Cells(1, 5).Value = "Nome do STA"
    wsE.Cells(1, 6).Value = "Itens em aberto"
    wsE.Cells(1, 7).Value = "Valor em aberto"
    wsE.Cells(1, 8).Value = T("Dias em aberto (m[a']x.)")
    wsE.Cells(1, COL_LOG).Value = T("[U']ltimo envio")

    lin = 1
    For Each k In OrdenarChaves(dic.Keys)
        cod = CStr(k)
        Set periodos = dic(cod)
        qtdTot = 0
        valTot = 0
        For Each k2 In periodos.Keys
            itens = periodos(k2)
            qtdTot = qtdTot + itens(0)
            valTot = valTot + itens(1)
        Next k2
        If qtdTot > 0 Then
            lin = lin + 1
            wsE.Cells(lin, 1).Value = dicEmail(cod)
            wsE.Cells(lin, 2).Value = T("Devolu[c][a~]o RPTG - ") & RegiaoDoSta(cod) & _
                                      Format$(qtdTot, "#,##0") & _
                                      T(" pe[c]a(s) pendente(s) - STA ") & cod
            wsE.Cells(lin, 3).Value = CorpoEmail(cod, CStr(dicNome(cod)), periodos, qtdTot, valTot)
            wsE.Cells(lin, 4).Value = cod
            wsE.Cells(lin, 5).Value = dicNome(cod)
            wsE.Cells(lin, 6).Value = qtdTot
            wsE.Cells(lin, 7).Value = valTot
            If dicDias.Exists(cod) Then wsE.Cells(lin, 8).Value = dicDias(cod)
            If dicLog.Exists(cod) Then wsE.Cells(lin, COL_LOG).Value = dicLog(cod)
        End If
    Next k

    Application.ScreenUpdating = True
    If silencioso Then
        Gravar "aba " & ABA_EMAILS & " reconstruida com " & (lin - 1) & " STA(s)"
    Else
        MsgBox "Aba " & ABA_EMAILS & T(" reconstru[i']da com ") & (lin - 1) & " STA(s) a cobrar.", vbInformation
    End If
End Sub

Private Function CorpoEmail(ByVal cod As String, ByVal nome As String, periodos As Object, _
                            ByVal qtdTot As Double, ByVal valTot As Double) As String
    Dim k As Variant, itens As Variant, s As String
    Dim ano As Long, mes As Long, prazo As Date, vencidas As Double

    s = "Prezado(a) parceiro(a) " & nome & " (" & cod & ")," & vbLf & vbLf & _
        T("Consta em nosso controle de devolu[c][a~]o RPTG ") & Format$(qtdTot, "#,##0") & _
        T(" pe[c]a(s) pendente(s) de devolu[c][a~]o, no valor total de R$ ") & _
        Format$(valTot, "#,##0.00") & "." & vbLf & vbLf & _
        T("Detalhamento por m[e^]s de refer[e^]ncia:") & vbLf

    For Each k In OrdenarChaves(periodos.Keys)
        itens = periodos(k)
        ano = Val(Left$(CStr(k), 4))
        mes = Val(Mid$(CStr(k), 6, 2))
        prazo = PrazoDevolucao(ano, mes)
        If prazoManual > 0 Then prazo = prazoManual
        s = s & Mid$(CStr(k), 6, 2) & "/" & Left$(CStr(k), 4) & ": " & _
            Format$(itens(0), "#,##0") & T(" pe[c]a(s) - R$ ") & Format$(itens(1), "#,##0.00") & _
            T(" - prazo at[e'] ") & Format$(prazo, "dd/mm/yyyy")
        If Date > prazo Then
            s = s & " (prazo vencido)"
            vencidas = vencidas + itens(1)
        End If
        s = s & vbLf
    Next k

    s = s & vbLf & T("Solicitamos a devolu[c][a~]o das pe[c]as dentro do prazo indicado acima.") & _
        T(" Ap[o']s a data limite, as pe[c]as n[a~]o devolvidas s[a~]o encaminhadas para d[e']bito do STA conforme contrato.")
    If vencidas > 0 Then
        s = s & T(" As pe[c]as com prazo vencido, no valor de R$ ") & Format$(vencidas, "#,##0.00") & _
            T(", ser[a~]o debitadas do STA conforme contrato.")
    End If
    s = s & vbLf & vbLf & _
        "Atenciosamente," & vbLf & _
        "Nome do Analista" & vbLf & _
        "Analista de Qualidade - RPTG" & vbLf & _
        T("S[a~]o Paulo - Brasil") & vbLf & _
        "E-mail: analista.rptg@exemplo.com.br"

    CorpoEmail = s
End Function

Private Function PeriodoCobravel(ByVal ano As Long, ByVal mes As Long) As Boolean
    If ano < 1900 Or mes < 1 Or mes > 12 Then Exit Function

    If mesesEscolhidos <> "" Then
        PeriodoCobravel = (InStr(1, mesesEscolhidos, ";" & Format$(ano, "0000") & "-" & _
                                 Format$(mes, "00") & ";") > 0)
        Exit Function
    End If

    If Not SO_MESES_NO_PRAZO Then
        PeriodoCobravel = True
        Exit Function
    End If
    PeriodoCobravel = (Date <= PrazoDevolucao(ano, mes))
End Function

Private Function PrazoDevolucao(ByVal ano As Long, ByVal mes As Long) As Date
    PrazoDevolucao = DateSerial(ano, mes + CB_MESES, CB_DIA_LIMITE)
End Function

Private Function RegiaoDoSta(ByVal cod As String) As String
    Dim reg As String, uf As String, partes As Variant

    If dicRegUf Is Nothing Then CarregarIsentas
    If dicRegUf Is Nothing Then Exit Function
    If Not dicRegUf.Exists(SoNumero(cod)) Then Exit Function

    partes = Split(CStr(dicRegUf(SoNumero(cod))), "|")
    reg = Trim$(CStr(partes(0)))
    If UBound(partes) >= 1 Then uf = Trim$(CStr(partes(1)))

    If reg = "" And uf = "" Then Exit Function
    If reg = "" Then
        RegiaoDoSta = UCase$(uf) & " - "
    ElseIf uf = "" Then
        RegiaoDoSta = reg & " - "
    Else
        RegiaoDoSta = reg & "/" & UCase$(uf) & " - "
    End If
End Function

Private Function CopiasDaSta(ByVal cod As String) As String
    Dim cc As String, item As Variant, reg As String, uf As String, partes As Variant

    If dicRegUf Is Nothing Then CarregarIsentas
    If Not dicRegUf Is Nothing Then
        If dicRegUf.Exists(SoNumero(cod)) Then
            partes = Split(CStr(dicRegUf(SoNumero(cod))), "|")
            reg = partes(0)
            If UBound(partes) >= 1 Then uf = partes(1)
        End If
    End If

    For Each item In Array(COPIA_PARA, EmailFinanceiro(), ConsultoresRegiao(reg, uf))
        If CStr(item) <> "" Then
            If InStr(1, cc, CStr(item), vbTextCompare) = 0 Then
                If cc <> "" Then cc = cc & "; "
                cc = cc & CStr(item)
            End If
        End If
    Next item
    CopiasDaSta = cc
End Function

Private Function T(ByVal s As String) As String
    s = Replace(s, "[c]", ChrW(231))
    s = Replace(s, "[C]", ChrW(199))
    s = Replace(s, "[a~]", ChrW(227))
    s = Replace(s, "[A~]", ChrW(195))
    s = Replace(s, "[a']", ChrW(225))
    s = Replace(s, "[a^]", ChrW(226))
    s = Replace(s, "[a`]", ChrW(224))
    s = Replace(s, "[e']", ChrW(233))
    s = Replace(s, "[e^]", ChrW(234))
    s = Replace(s, "[i']", ChrW(237))
    s = Replace(s, "[o']", ChrW(243))
    s = Replace(s, "[o~]", ChrW(245))
    s = Replace(s, "[u']", ChrW(250))
    s = Replace(s, "[U']", ChrW(218))
    T = s
End Function

Private Function EmailFinanceiro() As String
    On Error Resume Next
    EmailFinanceiro = Trim(CStr(ThisWorkbook.Worksheets(ABA_PARAM).Range(CEL_FINANCEIRO).Value))
    On Error GoTo 0
    If EmailFinanceiro = "" Then EmailFinanceiro = PARA_FIXO
End Function

' Cc da regiao: lido de Parametros O:P com fallback automatico no codigo
Private Function ConsultoresRegiao(ByVal regiao As String, ByVal uf As String) As String
    Dim ws As Worksheet, i As Long, ult As Long
    Dim rot As String, resto As String, alvoReg As String, alvoUf As String
    Dim coord As String, consultor As String

    alvoReg = Normalizar(regiao)
    alvoUf = Normalizar(uf)

    ' 1. Busca na aba Parametros (se existir preenchido)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_PARAM)
    On Error GoTo 0
    If Not ws Is Nothing And alvoReg <> "" Then
        ult = ws.Cells(ws.Rows.Count, PAR_COL_REG).End(xlUp).Row
        For i = 1 To ult
            rot = Normalizar(CStr(ws.Cells(i, PAR_COL_REG).Value))
            If rot <> "" Then
                If rot = alvoReg Then
                    coord = Trim(CStr(ws.Cells(i, PAR_COL_CC).Value))
                ElseIf Left$(rot, Len(alvoReg)) = alvoReg And alvoUf <> "" Then
                    resto = Mid$(rot, Len(alvoReg) + 1)
                    If InStr(1, resto, alvoUf) > 0 Then
                        consultor = Trim(CStr(ws.Cells(i, PAR_COL_CC).Value))
                    End If
                End If
            End If
        Next i
    End If

    ' 2. Fallback direto para o NORDESTE com o novo e-mail da Monique
    If consultor = "" And coord = "" Then
        Select Case UCase$(Trim$(uf))
            Case "BA", "PE", "CE", "MA", "PB", "RN", "AL", "SE", "PI"
                consultor = "consultor01@exemplo.com.br; consultor02@exemplo.com.br; consultor03@exemplo.com.br; consultor04@exemplo.com.br; consultor05@exemplo.com.br; consultor06@exemplo.com.br; consultor07@exemplo.com.br; consultor08@exemplo.com.br; consultor09@exemplo.com.br; consultor10@exemplo.com.br"
        End Select
    End If

    ConsultoresRegiao = JuntarEmails(consultor, coord)
End Function

Private Function JuntarEmails(ByVal a As String, ByVal b As String) As String
    Dim item As Variant, res As String, e As String
    For Each item In Split(Replace(a & ";" & b, ",", ";"), ";")
        e = Trim$(CStr(item))
        If e <> "" Then
            If InStr(1, res, e, vbTextCompare) = 0 Then
                If res <> "" Then res = res & "; "
                res = res & e
            End If
        End If
    Next item
    JuntarEmails = res
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

Public Sub Escolher_Banner_Devolucao()
    Dim dlg As Object, caminho As String
    Set dlg = Application.FileDialog(3)
    dlg.Title = "Selecione a imagem do banner da assinatura"
    dlg.AllowMultiSelect = False
    dlg.Filters.Clear
    dlg.Filters.Add "Imagens", "*.png; *.jpg; *.jpeg; *.gif"
    If dlg.Show <> -1 Then Exit Sub
    caminho = dlg.SelectedItems(1)
    GravarCaminhoBanner caminho
    MsgBox "Banner definido:" & vbCrLf & caminho, vbInformation, T("Devolu[c][a~]o RPTG")
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
    Dim p As String, resto As String, i As Long
    Dim base As Variant
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

Private Function MontarHtml(ByVal corpo As String, ByVal temBanner As Boolean) As String
    Dim html As String
    html = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f1f1f;"">" & _
           TextoParaHtml(corpo) & "</div>"
    If temBanner Then
        html = html & "<br><img src=""cid:bannerempresa"" width=""" & _
               LARGURA_BANNER & """ style=""border:0;display:block;"">"
    End If
    MontarHtml = "<html><body>" & html & "</body></html>"
End Function

Private Sub ProcessarEmails(ByVal enviar As Boolean, ByVal destinoTeste As String)
    Dim ws As Worksheet
    Dim olApp As Object, olMail As Object, olAnexo As Object
    Dim i As Long, ult As Long, n As Long, falhas As Long
    Dim para As String, assunto As String, corpo As String, copias As String
    Dim banner As String, temBanner As Boolean

    Set ws = AbaEmails()
    ult = UltimaLinha(ws)
    banner = CaminhoBanner()
    temBanner = (banner <> "")

    If Not temBanner And Not silencioso Then
        If MsgBox(T("N[a~]o encontrei o arquivo ") & ARQ_BANNER & " na pasta da planilha." & vbCrLf & _
                  "Dica: cancele e rode a macro Escolher_Banner_Devolucao para apontar a imagem." & vbCrLf & vbCrLf & _
                  "Continuar sem o banner na assinatura?", vbYesNo + vbExclamation, _
                  T("Devolu[c][a~]o RPTG")) <> vbYes Then Exit Sub
    End If

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        If silencioso Then
            Gravar "nao foi possivel abrir o Outlook nesta maquina"
        Else
            MsgBox T("N[a~]o foi poss[i']vel abrir o Outlook nesta m[a']quina."), vbCritical
        End If
        Exit Sub
    End If

    If Trim(CStr(ws.Cells(1, COL_LOG).Value)) = "" Then
        ws.Cells(1, COL_LOG).Value = T("[U']ltimo envio")
    End If

    Application.ScreenUpdating = False
    For i = 2 To ult
        If LinhaValida(ws, i) Then
            para = Trim(CStr(ws.Cells(i, 1).Value))
            assunto = CStr(ws.Cells(i, 2).Value)
            corpo = CStr(ws.Cells(i, 3).Value)
            copias = CopiasDaSta(CStr(ws.Cells(i, 4).Value))
            If destinoTeste <> "" Then
                para = destinoTeste
                assunto = "[TESTE] " & assunto
            End If

            On Error Resume Next
            Set olMail = olApp.CreateItem(0)
            With olMail
                .To = para
                If copias <> "" Then .cc = copias
                .Subject = assunto
                If temBanner Then
                    Set olAnexo = .Attachments.Add(banner, 1, 0)
                    olAnexo.PropertyAccessor.SetProperty _
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "bannerempresa"
                    Set olAnexo = Nothing
                End If
                .HTMLBody = MontarHtml(corpo, temBanner)
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
        Gravar "e-mails enviados: " & n & "; falhas: " & falhas
    Else
        If n = 0 And falhas = 0 Then
            MsgBox T("Nenhum e-mail foi gerado.") & vbCrLf & vbCrLf & Diagnostico(ws, ult), _
                   vbExclamation, T("Devolu[c][a~]o RPTG")
        Else
            MsgBox IIf(enviar, "E-mails enviados: ", T("E-mails abertos para revis[a~]o: ")) & n & _
                   IIf(falhas > 0, vbCrLf & "Falhas: " & falhas & " (veja a coluna I)", ""), vbInformation
        End If
    End If
End Sub

Private Function Diagnostico(ws As Worksheet, ByVal ult As Long) As String
    Dim i As Long, cod As String
    Dim linhas As Long, semEmail As Long, jaEnviados As Long
    Dim isentas As Long, semPeca As Long

    For i = 2 To ult
        linhas = linhas + 1
        If InStr(1, CStr(ws.Cells(i, 1).Value), "@") = 0 Then
            semEmail = semEmail + 1
        ElseIf PULAR_JA_ENVIADOS And Trim(CStr(ws.Cells(i, COL_LOG).Value)) <> "" Then
            jaEnviados = jaEnviados + 1
        Else
            cod = SoNumero(CStr(ws.Cells(i, 4).Value))
            If StaIsenta(cod) Then
                isentas = isentas + 1
            ElseIf Val(CStr(ws.Cells(i, 6).Value)) <= 0 Then
                semPeca = semPeca + 1
            End If
        End If
    Next i

    If linhas = 0 Then
        Diagnostico = T("A aba ") & ABA_EMAILS & T(" est[a'] vazia.") & vbCrLf & _
                      "Rode Alt+F8 > Atualizar_Emails_STA para montar a lista."
        Exit Function
    End If

    Diagnostico = T("Linhas na aba ") & ABA_EMAILS & ": " & linhas & vbCrLf & _
        "- sem e-mail na coluna A: " & semEmail & vbCrLf & _
        T("- j[a'] enviados (coluna I preenchida): ") & jaEnviados & vbCrLf & _
        "- MEI / Norte (isentos): " & isentas & vbCrLf & _
        T("- sem pe[c]as em aberto: ") & semPeca
End Function