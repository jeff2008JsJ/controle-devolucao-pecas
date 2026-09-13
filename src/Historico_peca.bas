Attribute VB_Name = "Historico_peca"
Option Explicit
'=====================================================================
' IMPORTACAO DO RELATORIO "HISTORICO DE PECAS" EXPORTADO DO SISTEMA DE CHAMADOS
'
' Fluxo:
'   1) Voce loga no sistema de chamados, marca "Exportar para Planilha?" e clica
'      em Consultar. O arquivo (Historico_de_pecas_v*.xls) cai na pasta
'      de Downloads.
'   2) Esta macro pega o arquivo MAIS RECENTE da pasta, le os dados,
'      atualiza as linhas que ja existem na Base_RPTG e acrescenta as
'      linhas novas (sem duplicar).
'   3) Depois roda Atualizar_Meses_Devidos / Atualizar_Aging.
'
' Colunas do export do sistema de chamados (confirmadas no arquivo de exemplo):
'   Chamado | Data de Encerramento | Contrato | Garantia de Fabrica |
'   STA | Fornecedor | Cod. Familia | Familia | Cod. Peca | Peca |
'   Quantidade | Data de Fabricacao | Situacao RPTG | Data de Analise |
'   Defeito STA | Defeito MTF | Status | Destino da Peca | Observacao |
'   Auditor
'
' O que o sistema de chamados NAO traz e a macro completa sozinha:
'   Codigo_STA  -> tirado do proprio texto da STA, ex. "... (58578)"
'   Mes / Ano   -> da Data de Encerramento
'   Valor_Debito_Unit / Tipo / Descricao -> aba Base_Valores (Cod. Peca)
'   UF / Coordenador / MEI               -> aba STAs (Cod STA)
'   NF / Transportadora                  -> ficam em branco (voce preenche)
'
' Rotinas para usar:
'   Importar_Chamados            -> manual, mostra o resumo na tela
'   Importar_Chamados_Automatico -> silenciosa (Agendador de Tarefas)
'   Escolher_Pasta_Chamados      -> troca a pasta onde procurar o arquivo
'=====================================================================
Private Const IM_ABA_BASE    As String = "Base_RPTG"
Private Const IM_TABELA      As String = "tblBase"
Private Const IM_ABA_PARAM   As String = "Parametros"
Private Const IM_ABA_VALORES As String = "Base_Valores"
Private Const IM_ABA_STAS    As String = "STAs"
Private Const IM_CEL_PASTA   As String = "P18"   ' pasta onde procurar o export
Private Const IM_CEL_ULTIMO  As String = "P19"   ' ultimo arquivo importado (controle)
Private Const IM_EXTENSOES   As String = "xls|xlsx|xlsm|xlsb|csv|txt"
Private Const IM_HORAS_MAX   As Long = 0         ' 0 = sem limite de idade do arquivo
Private Const IM_ARQ_LOG     As String = "Log_Importacao_Chamados.txt"
' False = importa e NAO dispara o recalculo pesado (rode depois, quando quiser,
'         Atualizar_Meses_Devidos). Em planilhas grandes deixe False.
Private Const IM_RECALCULAR   As Boolean = False
' status inicial das linhas novas - montado por codigo de caractere para
' nao depender de acentos ao colar o modulo no VBA ("Aguardando devolucao")
Private Function IM_STATUS_INI() As String
    IM_STATUS_INI = "Aguardando devolu" & ChrW$(231) & ChrW$(227) & "o"
End Function
'---------------------------------------------------------------------
' 1) EXECUCAO MANUAL
'---------------------------------------------------------------------
Public Sub Importar_Chamados()
    Dim msg As String
    msg = IM_Executar(False)
    IM_Log msg
    MsgBox msg, vbInformation, "Importar sistema de chamados"
End Sub
'---------------------------------------------------------------------
' 2) EXECUCAO AUTOMATICA (Agendador de Tarefas / .vbs)
'---------------------------------------------------------------------
Public Sub Importar_Chamados_Automatico()
    Dim msg As String
    On Error Resume Next
    msg = IM_Executar(True)
    IM_Log msg
End Sub
'---------------------------------------------------------------------
' 3) ESCOLHER A PASTA DO EXPORT (por padrao usa a pasta Downloads)
'---------------------------------------------------------------------
Public Sub Escolher_Pasta_Chamados()
    Dim fd As Object, ws As Worksheet
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = "Selecione a pasta onde o sistema de chamados salva o arquivo exportado"
    fd.InitialFileName = IM_PastaBusca() & "\"
    If fd.Show <> -1 Then Exit Sub
    Set ws = IM_Param()
    If ws Is Nothing Then
        MsgBox "Aba " & IM_ABA_PARAM & " nao encontrada.", vbExclamation
        Exit Sub
    End If
    ws.Range(IM_CEL_PASTA).Value = fd.SelectedItems(1)
    MsgBox "Pasta gravada:" & vbCrLf & fd.SelectedItems(1), vbInformation
End Sub
'=====================================================================
' NUCLEO DA IMPORTACAO
'=====================================================================
Private Function IM_Executar(ByVal silencioso As Boolean) As String
    Dim pasta As String, arq As String
    Dim wbOrig As Workbook, wsOrig As Worksheet
    Dim ws As Worksheet, lo As ListObject
    Dim mapa As Object, chaves As Object, dicVal As Object, dicSta As Object
    Dim linCab As Long, ultLinOrig As Long, ultColOrig As Long
    Dim i As Long, c As Long, k As String
    Dim nLidas As Long, nNovas As Long, nAtu As Long, nIguais As Long, nIgnoradas As Long
    Dim novas() As Variant, nNovasBuf As Long, nColBase As Long
    Dim calcAnt As XlCalculation, telaAnt As Boolean, eventosAnt As Boolean
    Dim diagCols As String
    On Error GoTo trata
    telaAnt = Application.ScreenUpdating
    eventosAnt = Application.EnableEvents
    calcAnt = Application.Calculation
    Set ws = ThisWorkbook.Worksheets(IM_ABA_BASE)
    Set lo = ws.ListObjects(IM_TABELA)
    nColBase = lo.ListColumns.Count
    pasta = IM_PastaBusca()
    arq = IM_ArquivoMaisRecente(pasta)
    If arq = "" Then
        IM_Executar = "Nenhum arquivo exportado encontrado em:" & vbCrLf & pasta & _
                      vbCrLf & vbCrLf & "Exporte o relatorio no sistema de chamados e rode de novo."
        Exit Function
    End If
    ' evita reimportar o mesmo arquivo na execucao automatica
    If silencioso Then
        If IM_Assinatura(arq) = IM_UltimoImportado() Then
            IM_Executar = "Arquivo ja importado antes (nada a fazer): " & arq
            Exit Function
        End If
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Set wbOrig = Workbooks.Open(Filename:=arq, ReadOnly:=True, UpdateLinks:=0, Local:=True)
    Set wsOrig = wbOrig.Worksheets(1)
    linCab = IM_LinhaCabecalho(wsOrig)
    If linCab = 0 Then
        wbOrig.Close SaveChanges:=False
        GoTo fim_semmapa
    End If
    ultColOrig = wsOrig.Cells(linCab, wsOrig.Columns.Count).End(xlToLeft).Column
    If ultColOrig < 5 Then ultColOrig = 30
    ultLinOrig = linCab
    For c = 1 To ultColOrig
        If wsOrig.Cells(wsOrig.Rows.Count, c).End(xlUp).Row > ultLinOrig Then
            ultLinOrig = wsOrig.Cells(wsOrig.Rows.Count, c).End(xlUp).Row
        End If
    Next c
    Set mapa = IM_MapearColunas(wsOrig, linCab, ultColOrig)
    diagCols = Join(mapa.Keys, ", ")
    If Not (mapa.Exists("Chamado") And mapa.Exists("Codigo_Peca")) _
       Or mapa.Count < 4 Or ultLinOrig <= linCab Then
        wbOrig.Close SaveChanges:=False
        GoTo fim_semmapa
    End If
    ' ---- le o arquivo de origem de uma vez
    Dim orig As Variant
    orig = wsOrig.Range(wsOrig.Cells(linCab + 1, 1), wsOrig.Cells(ultLinOrig, ultColOrig)).Value
    wbOrig.Close SaveChanges:=False
    Set wbOrig = Nothing
    ' ---- le a base e monta o indice de chaves
    Dim base As Variant, primLin As Long, nLinBase As Long
    primLin = lo.Range.Row + 1
    nLinBase = lo.ListRows.Count
    If nLinBase > 0 Then
        base = ws.Range(ws.Cells(primLin, lo.Range.Column), _
                        ws.Cells(primLin + nLinBase - 1, lo.Range.Column + nColBase - 1)).Value
    End If
    Dim cCham As Long, cCodSta As Long, cPeca As Long, cSta As Long
    Dim cMes As Long, cAno As Long, cVal As Long, cTipo As Long, cDesc As Long
    Dim cUF As Long, cCoord As Long, cMEI As Long, cEnc As Long, cStatus As Long
    cCham = IM_Col(lo, "Chamado")
    cCodSta = IM_Col(lo, "Codigo_STA")
    cPeca = IM_Col(lo, "Codigo_Peca")
    cSta = IM_Col(lo, "STA")
    cMes = IM_Col(lo, "Mes")
    cAno = IM_Col(lo, "Ano")
    cVal = IM_Col(lo, "Valor_Debito_Unit")
    cTipo = IM_Col(lo, "Tipo")
    cDesc = IM_Col(lo, "Descricao")
    cUF = IM_Col(lo, "UF")
    cCoord = IM_Col(lo, "Coordenador")
    cMEI = IM_Col(lo, "MEI")
    cEnc = IM_Col(lo, "Dta_Encerramento")
    cStatus = IM_Col(lo, "Status")
    Set chaves = CreateObject("Scripting.Dictionary")
    For i = 1 To nLinBase
        k = IM_Chave(base(i, cCham), base(i, cCodSta), base(i, cPeca))
        If k <> "" Then
            If Not chaves.Exists(k) Then chaves.Add k, i
        End If
    Next i
    Set dicVal = IM_DicValores()
    Set dicSta = IM_DicSTAs()
    ' ---- colunas que NAO devem ser escritas (formulas da tabela)
    Dim ehFormula() As Boolean
    ReDim ehFormula(1 To nColBase)
    For c = 1 To nColBase
        If nLinBase > 0 Then
            ehFormula(c) = ws.Cells(primLin + nLinBase - 1, lo.Range.Column + c - 1).HasFormula
        End If
    Next c
    ' ---- campos atualizaveis nas linhas que ja existem
    Dim atualizaveis As Variant, na As Long, colAtu() As Long
    atualizaveis = Array("Status", "Data_Validacao", "QTD", "Descricao", _
                         "Dta_Encerramento", "STA", "Contrato")
    ReDim colAtu(LBound(atualizaveis) To UBound(atualizaveis))
    For na = LBound(atualizaveis) To UBound(atualizaveis)
        colAtu(na) = IM_Col(lo, CStr(atualizaveis(na)))
    Next na
    Dim nLinOrig As Long, nColLidas As Long
    nLinOrig = UBound(orig, 1)
    nColLidas = UBound(orig, 2)
    ReDim novas(1 To nLinOrig, 1 To nColBase)
    nNovasBuf = 0
    Dim vDest As Variant, colDest As Long, linBase As Long, alterou As Boolean
    Dim vCham As Variant, vPeca As Variant, vStaTxt As String, sCodSta As String
    Dim dtEnc As Variant, ar As Variant
    For i = 1 To nLinOrig
        vCham = IM_ValorTratado(orig, i, mapa, "Chamado", nColLidas)
        vPeca = IM_ValorTratado(orig, i, mapa, "Codigo_Peca", nColLidas)
        vStaTxt = Trim$(CStr(IM_Valor(orig, i, mapa, "STA", nColLidas) & ""))
        sCodSta = Trim$(CStr(IM_ValorTratado(orig, i, mapa, "Codigo_STA", nColLidas) & ""))
        If sCodSta = "" Then sCodSta = IM_CodSTA(vStaTxt)
        dtEnc = IM_ParaData(IM_Valor(orig, i, mapa, "Dta_Encerramento", nColLidas))
        If IM_Vazio(vCham) Or IM_Vazio(vPeca) Then
            nIgnoradas = nIgnoradas + 1          ' linha em branco / rodape
            GoTo prox_linha
        End If
        nLidas = nLidas + 1
        k = IM_Chave(vCham, sCodSta, vPeca)
        If k <> "" And chaves.Exists(k) Then
            linBase = chaves(k)
            If linBase = 0 Then
                nIgnoradas = nIgnoradas + 1      ' chave repetida no proprio export
                nLidas = nLidas - 1
                GoTo prox_linha
            End If
            '---- ja existe: atualiza somente o que veio preenchido e mudou
            alterou = False
            For na = LBound(atualizaveis) To UBound(atualizaveis)
                colDest = colAtu(na)
                If colDest > 0 Then
                    If Not ehFormula(colDest) Then
                        vDest = IM_ValorTratado(orig, i, mapa, CStr(atualizaveis(na)), nColLidas)
                        If Not IM_Vazio(vDest) Then
                            If IM_Difere(base(linBase, colDest), vDest) Then
                                If IM_PodeAtualizar(CStr(atualizaveis(na)), base(linBase, colDest), vDest) Then
                                    ws.Cells(primLin + linBase - 1, lo.Range.Column + colDest - 1).Value = vDest
                                    base(linBase, colDest) = vDest
                                    alterou = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next na
            If alterou Then nAtu = nAtu + 1 Else nIguais = nIguais + 1
        Else
            '---- linha nova
            nNovasBuf = nNovasBuf + 1
            For c = 1 To nColBase
                If Not ehFormula(c) Then
                    novas(nNovasBuf, c) = IM_ValorTratado(orig, i, mapa, _
                                             lo.ListColumns(c).Name, nColLidas)
                End If
            Next c
            ' completa o que o sistema de chamados nao traz
            If cCodSta > 0 And sCodSta <> "" Then novas(nNovasBuf, cCodSta) = CDbl(sCodSta)
            If cEnc > 0 And Not IM_Vazio(dtEnc) Then novas(nNovasBuf, cEnc) = dtEnc
            If IsDate(dtEnc) Then
                If cMes > 0 Then novas(nNovasBuf, cMes) = Month(CDate(dtEnc))
                If cAno > 0 Then novas(nNovasBuf, cAno) = Year(CDate(dtEnc))
            End If
            ' valor / tipo / descricao pela Base_Valores
            If dicVal.Exists(UCase$(Trim$(CStr(vPeca)))) Then
                ar = dicVal(UCase$(Trim$(CStr(vPeca))))
                If cVal > 0 Then If IM_Vazio(novas(nNovasBuf, cVal)) Then novas(nNovasBuf, cVal) = ar(0)
                If cTipo > 0 Then If IM_Vazio(novas(nNovasBuf, cTipo)) Then novas(nNovasBuf, cTipo) = ar(1)
                If cDesc > 0 Then If IM_Vazio(novas(nNovasBuf, cDesc)) Then novas(nNovasBuf, cDesc) = ar(2)
            End If
            ' UF / coordenador / MEI pela aba STAs
            If dicSta.Exists(sCodSta) Then
                ar = dicSta(sCodSta)
                If cUF > 0 Then If IM_Vazio(novas(nNovasBuf, cUF)) Then novas(nNovasBuf, cUF) = ar(0)
                If cCoord > 0 Then If IM_Vazio(novas(nNovasBuf, cCoord)) Then novas(nNovasBuf, cCoord) = ar(1)
                If cMEI > 0 Then If IM_Vazio(novas(nNovasBuf, cMEI)) Then novas(nNovasBuf, cMEI) = ar(2)
            End If
            If cStatus > 0 Then
                If IM_Vazio(novas(nNovasBuf, cStatus)) Then novas(nNovasBuf, cStatus) = IM_STATUS_INI
            End If
            If k <> "" Then chaves.Add k, 0
            nNovas = nNovas + 1
        End If
prox_linha:
    Next i
    ' ---- grava as linhas novas em bloco e estende a tabela
    If nNovasBuf > 0 Then
        Dim linIni As Long, linFim As Long
        linIni = primLin + nLinBase
        linFim = linIni + nNovasBuf - 1
        ws.Range(ws.Cells(linIni, lo.Range.Column), _
                 ws.Cells(linFim, lo.Range.Column + nColBase - 1)).Value = novas
        lo.Resize ws.Range(ws.Cells(lo.Range.Row, lo.Range.Column), _
                           ws.Cells(linFim, lo.Range.Column + nColBase - 1))
        ' repete as formulas da tabela nas linhas novas
        For c = 1 To nColBase
            If ehFormula(c) Then
                ws.Range(ws.Cells(linIni - 1, lo.Range.Column + c - 1), _
                         ws.Cells(linFim, lo.Range.Column + c - 1)).FillDown
            End If
        Next c
        IM_FormatarData ws, lo, "Dta_Encerramento", linIni, linFim
        IM_FormatarData ws, lo, "Data_Validacao", linIni, linFim
    End If
    IM_GravarUltimo IM_Assinatura(arq)
    ' ---- recalculos da planilha (opcionais - ver IM_RECALCULAR no topo)
    Dim posRecalc As String
    If IM_RECALCULAR Then
        Application.Calculation = calcAnt
        Application.Calculate
        On Error Resume Next
        Err.Clear
        Application.Run "'" & ThisWorkbook.Name & "'!Atualizar_Meses_Devidos"
        If Err.Number <> 0 Then
            Err.Clear
            Application.Run "'" & ThisWorkbook.Name & "'!Atualizar_Aging"
            If Err.Number <> 0 Then
                posRecalc = vbCrLf & "(atencao: nao encontrei Atualizar_Meses_Devidos/Atualizar_Aging)"
                Err.Clear
            End If
        End If
        On Error GoTo trata
    Else
        Application.Calculation = calcAnt
        posRecalc = vbCrLf & "(rode Atualizar_Meses_Devidos quando quiser atualizar aging/resumos)"
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    IM_Executar = "Arquivo: " & arq & vbCrLf & _
                  "Linhas lidas: " & nLidas & vbCrLf & _
                  "Linhas novas acrescentadas: " & nNovas & vbCrLf & _
                  "Linhas atualizadas: " & nAtu & vbCrLf & _
                  "Linhas sem mudanca: " & nIguais & vbCrLf & _
                  "Linhas ignoradas (em branco): " & nIgnoradas & vbCrLf & _
                  "Colunas reconhecidas: " & diagCols & posRecalc
    Exit Function
fim_semmapa:
    IM_Restaurar telaAnt, eventosAnt, calcAnt
    IM_Executar = "Nao reconheci as colunas do arquivo:" & vbCrLf & arq & _
                  vbCrLf & "Colunas reconhecidas: " & diagCols & _
                  vbCrLf & "Confirme se e o export do Historico de Pecas."
    Exit Function
trata:
    On Error Resume Next
    If Not wbOrig Is Nothing Then wbOrig.Close SaveChanges:=False
    IM_Restaurar telaAnt, eventosAnt, calcAnt
    IM_Executar = "Erro na importacao: " & Err.Number & " - " & Err.Description
End Function
Private Sub IM_Restaurar(ByVal telaAnt As Boolean, ByVal eventosAnt As Boolean, _
                         ByVal calcAnt As XlCalculation)
    On Error Resume Next
    Application.Calculation = calcAnt
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
' socorro: destrava a tela se alguma macro parou no meio
Public Sub Destravar_Tela()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Tela liberada.", vbInformation
End Sub
' nao deixa o import "voltar" um status que voce ja evoluiu na mao
Private Function IM_PodeAtualizar(ByVal campo As String, ByVal atual As Variant, _
                                  ByVal novo As Variant) As Boolean
    IM_PodeAtualizar = True
    If StrComp(campo, "Status", vbTextCompare) = 0 Then
        If Not IM_Vazio(atual) Then
            If StrComp(Trim$(CStr(novo)), IM_STATUS_INI, vbTextCompare) = 0 Then
                IM_PodeAtualizar = False
            End If
        End If
    End If
End Function
'=====================================================================
' DICIONARIOS AUXILIARES
'=====================================================================
' Base_Valores: Cod. Peca (A) -> Valor Unitario (C), Tipo (H), Peca (B)
Private Function IM_DicValores() As Object
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(IM_ABA_VALORES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set IM_DicValores = d
        Exit Function
    End If
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        k = UCase$(Trim$(CStr(ws.Cells(i, 1).Value)))
        If k <> "" And Not d.Exists(k) Then
            d.Add k, Array(ws.Cells(i, 3).Value, ws.Cells(i, 8).Value, ws.Cells(i, 2).Value)
        End If
    Next i
    Set IM_DicValores = d
End Function
' STAs: Cod STA (F) -> UF (A), Coordenador (E), MEI? (J)
Private Function IM_DicSTAs() As Object
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(IM_ABA_STAS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set IM_DicSTAs = d
        Exit Function
    End If
    ult = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For i = 2 To ult
        k = Trim$(CStr(ws.Cells(i, 6).Value))
        If IsNumeric(k) And k <> "" Then k = CStr(CLng(CDbl(k)))
        If k <> "" And Not d.Exists(k) Then
            d.Add k, Array(ws.Cells(i, 1).Value, ws.Cells(i, 5).Value, ws.Cells(i, 10).Value)
        End If
    Next i
    Set IM_DicSTAs = d
End Function
' tira o codigo do texto da STA: "FRIOTEC ... (58578)" -> "58578"
Private Function IM_CodSTA(ByVal s As String) As String
    Dim p1 As Long, p2 As Long, dentro As String
    s = Trim$(s)
    p2 = InStrRev(s, ")")
    If p2 = 0 Then Exit Function
    p1 = InStrRev(s, "(", p2)
    If p1 = 0 Then Exit Function
    dentro = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
    If dentro <> "" And IsNumeric(dentro) Then IM_CodSTA = CStr(CLng(CDbl(dentro)))
End Function
'=====================================================================
' ARQUIVO / PASTA
'=====================================================================
Private Function IM_PastaBusca() As String
    Dim ws As Worksheet, p As String
    Set ws = IM_Param()
    If Not ws Is Nothing Then p = Trim$(CStr(ws.Range(IM_CEL_PASTA).Value))
    If p = "" Then p = Environ$("USERPROFILE") & "\Downloads"
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    IM_PastaBusca = p
End Function
Private Function IM_ArquivoMaisRecente(ByVal pasta As String) As String
    Dim fso As Object, f As Object, ext As String, exts As Variant, i As Long
    Dim melhor As String, melhorData As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pasta) Then Exit Function
    exts = Split(LCase$(IM_EXTENSOES), "|")
    For Each f In fso.GetFolder(pasta).Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If Left$(f.Name, 2) <> "~$" Then
            For i = LBound(exts) To UBound(exts)
                If ext = exts(i) Then
                    If IM_HORAS_MAX = 0 Or f.DateLastModified >= Now - IM_HORAS_MAX / 24 Then
                        If melhor = "" Or f.DateLastModified > melhorData Then
                            melhor = f.Path
                            melhorData = f.DateLastModified
                        End If
                    End If
                    Exit For
                End If
            Next i
        End If
    Next f
    IM_ArquivoMaisRecente = melhor
End Function
Private Function IM_Assinatura(ByVal arq As String) As String
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    IM_Assinatura = arq & "|" & Format$(fso.GetFile(arq).DateLastModified, "yyyy-mm-dd hh:nn:ss") & _
                    "|" & fso.GetFile(arq).Size
End Function
Private Function IM_UltimoImportado() As String
    Dim ws As Worksheet
    Set ws = IM_Param()
    If Not ws Is Nothing Then IM_UltimoImportado = Trim$(CStr(ws.Range(IM_CEL_ULTIMO).Value))
End Function
Private Sub IM_GravarUltimo(ByVal s As String)
    Dim ws As Worksheet
    Set ws = IM_Param()
    If Not ws Is Nothing Then ws.Range(IM_CEL_ULTIMO).Value = "'" & s
End Sub
Private Function IM_Param() As Worksheet
    On Error Resume Next
    Set IM_Param = ThisWorkbook.Worksheets(IM_ABA_PARAM)
End Function
Private Sub IM_Log(ByVal msg As String)
    Dim fso As Object, ts As Object, cam As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    cam = ThisWorkbook.Path & "\" & IM_ARQ_LOG
    Set ts = fso.OpenTextFile(cam, 8, True)
    ts.WriteLine Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & Replace(msg, vbCrLf, " | ")
    ts.Close
End Sub
'=====================================================================
' CABECALHOS
'=====================================================================
Private Function IM_LinhaCabecalho(ByVal wsOrig As Worksheet) As Long
    Dim r As Long, c As Long, achou As Long, nome As String
    For r = 1 To 20
        achou = 0
        For c = 1 To 40
            nome = IM_Destino(IM_Norm(CStr(wsOrig.Cells(r, c).Value)))
            If nome <> "" Then achou = achou + 1
        Next c
        If achou >= 4 Then
            IM_LinhaCabecalho = r
            Exit Function
        End If
    Next r
End Function
' devolve dicionario: nome da coluna da Base -> numero da coluna no arquivo
Private Function IM_MapearColunas(ByVal wsOrig As Worksheet, ByVal linCab As Long, _
                                  ByVal ultCol As Long) As Object
    Dim d As Object, c As Long, nome As String
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To ultCol
        nome = IM_Destino(IM_Norm(CStr(wsOrig.Cells(linCab, c).Value)))
        If nome <> "" Then
            If Not d.Exists(nome) Then d.Add nome, c
        End If
    Next c
    Set IM_MapearColunas = d
End Function
' normaliza texto: minusculo, sem acento, sem espacos/pontuacao
' os acentos sao tratados pelo CODIGO do caractere (AscW), assim o de-para
' dos cabecalhos funciona mesmo se o modulo for colado com acentos tortos
Private Function IM_Norm(ByVal s As String) As String
    Dim i As Long, cod As Long, ch As String, r As String
    s = Trim$(s)
    For i = 1 To Len(s)
        cod = AscW(Mid$(s, i, 1))
        If cod < 0 Then cod = cod + 65536
        Select Case cod
            Case 192 To 197, 224 To 229: ch = "a"      ' A a com acento
            Case 199, 231:               ch = "c"      ' C cedilha
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
    IM_Norm = r
End Function
' de-para: cabecalho do sistema de chamados (normalizado) -> coluna da Base_RPTG
Private Function IM_Destino(ByVal h As String) As String
    Select Case h
        Case "chamado", "chmado", "nrchamado", "numerochamado", "nchamado", "numchamado"
            IM_Destino = "Chamado"
        Case "codigosta", "codsta", "cdsta", "codigodasta"
            IM_Destino = "Codigo_STA"
        Case "sta", "nomesta", "razaosocial", "nomedasta", "prestador", "assistencia"
            IM_Destino = "STA"
        Case "codpeca", "codigopeca", "codigodapeca", "codigoitem", "codigo", _
             "cdpea", "cdpeca", "codpea", "cdigopea", "cdigopeca"
            IM_Destino = "Codigo_Peca"
        Case "peca", "pea", "descricao", "descricaopeca", "descpeca", "descricaodapeca", _
             "descricaoitem", "item"
            IM_Destino = "Descricao"
        Case "quantidade", "qtd", "qtde", "qtdpecas", "qt"
            IM_Destino = "QTD"
        Case "valordebito", "valordebitounit", "valorunitario", "valorunit", _
             "vlrdebito", "valordebitounitario", "valorpeca"
            IM_Destino = "Valor_Debito_Unit"
        Case "dtaencerramento", "dataencerramento", "dtencerramento", "dataencerr", _
             "dtencerr", "datadeencerramento", "encerramento"
            IM_Destino = "Dta_Encerramento"
        Case "datavalidacao", "dtavalidacao", "datadevalidacao", "dataanalise", _
             "datadeanalise", "dtanalise", "validacao"
            IM_Destino = "Data_Validacao"
        Case "status", "statuspeca"
            IM_Destino = "Status"
        Case "nf", "notafiscal", "numeronf", "nrnf", "nnf", "notafiscaldevolucao"
            IM_Destino = "NF"
        Case "mei"
            IM_Destino = "MEI"
        Case "uf", "ufsta"
            IM_Destino = "UF"
        Case "coordenador", "coordenacao"
            IM_Destino = "Coordenador"
        Case "contrato", "cliente"
            IM_Destino = "Contrato"
        Case "tipo", "tipoequipamento"
            IM_Destino = "Tipo"
        Case "mes"
            IM_Destino = "Mes"
        Case "ano"
            IM_Destino = "Ano"
        Case "transportadora", "transp", "transportador"
            IM_Destino = "Transportadora"
        Case Else
            ' Familia, Cod. Familia, Situacao RPTG, Defeito STA, Defeito MTF,
            ' Destino da Peca, Observacao, Auditor, Fornecedor, Garantia de
            ' Fabrica e Data de Fabricacao nao sao importados.
            IM_Destino = ""
    End Select
End Function
Private Function IM_Col(ByVal lo As ListObject, ByVal nome As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, nome, vbTextCompare) = 0 Then
            IM_Col = i
            Exit Function
        End If
    Next i
End Function
'=====================================================================
' VALORES
'=====================================================================
Private Function IM_Valor(ByRef orig As Variant, ByVal lin As Long, ByVal mapa As Object, _
                          ByVal campo As String, ByVal nCol As Long) As Variant
    If Not mapa.Exists(campo) Then Exit Function
    If mapa(campo) > nCol Then Exit Function
    IM_Valor = orig(lin, mapa(campo))
End Function
' valor ja convertido para o tipo da coluna de destino
Private Function IM_ValorTratado(ByRef orig As Variant, ByVal lin As Long, ByVal mapa As Object, _
                                 ByVal campo As String, ByVal nCol As Long) As Variant
    Dim v As Variant
    v = IM_Valor(orig, lin, mapa, campo, nCol)
    If IM_Vazio(v) Then Exit Function
    Select Case campo
        Case "Dta_Encerramento", "Data_Validacao"
            IM_ValorTratado = IM_ParaData(v)
        Case "Chamado", "Codigo_STA", "QTD", "Mes", "Ano", "Valor_Debito_Unit"
            IM_ValorTratado = IM_ParaNumero(v)
        Case "NF", "Codigo_Peca"
            IM_ValorTratado = Trim$(Replace(CStr(v), "  ", " "))
        Case Else
            If VarType(v) = vbString Then
                IM_ValorTratado = Trim$(CStr(v))
            Else
                IM_ValorTratado = v
            End If
    End Select
End Function
Private Function IM_Vazio(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IM_Vazio = True
    ElseIf IsNull(v) Then
        IM_Vazio = True
    ElseIf VarType(v) = vbString Then
        IM_Vazio = (Trim$(CStr(v)) = "")
    End If
End Function
Private Function IM_ParaData(ByVal v As Variant) As Variant
    Dim s As String, p As Variant, d As Long, m As Long, a As Long
    If IM_Vazio(v) Then Exit Function
    If VarType(v) = vbDate Then
        IM_ParaData = CDate(v)
        Exit Function
    End If
    If IsNumeric(v) And VarType(v) <> vbString Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            IM_ParaData = CDate(CDbl(v))
            Exit Function
        End If
    End If
    s = Trim$(CStr(v))
    s = Replace(Replace(s, "-", "/"), ".", "/")
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    p = Split(s, "/")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            d = CLng(p(0))
            m = CLng(p(1))
            a = CLng(p(2))
            If a > 9999 Then a = CLng(Left$(CStr(a), 4))   ' corrige 20256 -> 2025
            If a < 100 Then a = 2000 + a
            If d >= 1 And d <= 31 And m >= 1 And m <= 12 Then
                IM_ParaData = DateSerial(a, m, d)
                Exit Function
            End If
        End If
    End If
    If IsDate(s) Then IM_ParaData = CDate(s) Else IM_ParaData = v
End Function
Private Function IM_ParaNumero(ByVal v As Variant) As Variant
    Dim s As String
    If IM_Vazio(v) Then Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        IM_ParaNumero = v
        Exit Function
    End If
    s = Trim$(CStr(v))
    s = Replace(s, "R$", "")
    s = Replace(s, " ", "")
    If InStr(s, ",") > 0 Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", ".")
    End If
    If IsNumeric(s) Then
        IM_ParaNumero = CDbl(s)
    Else
        IM_ParaNumero = v
    End If
End Function
Private Function IM_Difere(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) = vbDate And VarType(b) = vbDate Then
        IM_Difere = (Int(CDbl(CDate(a))) <> Int(CDbl(CDate(b))))
    ElseIf IsNumeric(a) And IsNumeric(b) And Not IM_Vazio(a) And Not IM_Vazio(b) Then
        IM_Difere = (Abs(CDbl(a) - CDbl(b)) > 0.0001)
    Else
        IM_Difere = (StrComp(Trim$(CStr(a & "")), Trim$(CStr(b & "")), vbTextCompare) <> 0)
    End If
End Function
' chave de comparacao: chamado + codigo da STA + codigo da peca
Private Function IM_Chave(ByVal cham As Variant, ByVal sta As Variant, ByVal peca As Variant) As String
    Dim sCham As String, sSta As String, sPeca As String
    sCham = IM_TxtNum(cham)
    sSta = IM_TxtNum(sta)
    sPeca = UCase$(Trim$(CStr(peca & "")))
    If sCham = "" Or sPeca = "" Then Exit Function
    IM_Chave = sCham & "|" & sSta & "|" & sPeca
End Function
Private Function IM_TxtNum(ByVal v As Variant) As String
    If IM_Vazio(v) Then Exit Function
    If IsNumeric(v) Then
        IM_TxtNum = Format$(CDbl(v), "0.####")
    Else
        IM_TxtNum = Trim$(CStr(v))
    End If
End Function
Private Sub IM_FormatarData(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal nome As String, _
                            ByVal linIni As Long, ByVal linFim As Long)
    Dim c As Long
    c = IM_Col(lo, nome)
    If c = 0 Then Exit Sub
    ws.Range(ws.Cells(linIni, lo.Range.Column + c - 1), _
             ws.Cells(linFim, lo.Range.Column + c - 1)).NumberFormat = "dd/mm/yyyy"
End Sub
'=====================================================================
' IMPORTAR + ENVIAR NO MESMO PASSO (para o Agendador mensal)
'=====================================================================
Public Sub Rotina_Mensal_Chamados_e_Emails()
    Importar_Chamados_Automatico
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!Envio_Mensal_Automatico_Regiao"
End Sub