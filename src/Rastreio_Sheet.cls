Attribute VB_Name = "Planilha1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Const URL_BASE As String = _
    "https://www.braspress.com.br/site/w/tracking/search?documentType=NOTAFISCAL"

' colunas da aba Rastreio
Private Const C_CHAVE As Long = 1
Private Const C_TRANSP As Long = 2
Private Const C_NF As Long = 3
Private Const C_CODSTA As Long = 4
Private Const C_STA As Long = 5
Private Const C_ITENS As Long = 6
Private Const C_VALOR As Long = 7
Private Const C_STATUS As Long = 8
Private Const C_OCORR As Long = 9
Private Const C_DTOCORR As Long = 10
Private Const C_PREVISAO As Long = 11
Private Const C_ENTREGA As Long = 12
Private Const C_CTE As Long = 13
Private Const C_CONSULTA As Long = 14
Private Const C_LINK As Long = 15

' status da Base_RPTG considerados "em transporte" (minúsculas, entre barras)
Private Const EM_TRANSPORTE As String = "|em transporte|aguardando transporte|"


' ==========================================================================
'  Rotina do dia a dia
' ==========================================================================
Public Sub Rastreio_Atualizar_Tudo()
    Rastreio_Reconstruir_Lista
    Atualizar_Rastreio_Braspress
End Sub


' ==========================================================================
'  Refaz a aba Rastreio a partir da Base_RPTG, preservando o que já foi
'  consultado para as combinações Transportadora+NF que continuam na lista.
' ==========================================================================
Public Sub Rastreio_Reconstruir_Lista()
    Dim wsB As Worksheet, wsR As Worksheet
    Dim cab As Variant, dadosBase As Variant, colunas As Variant
    Dim guardado As Variant, chavesOrd As Variant
    Dim iStatus As Long, iTransp As Long, iNF As Long, iCod As Long, iSta As Long
    Dim chaves As Object, cods As Object, nomes As Object, antigos As Object
    Dim i As Long, k As Long, ult As Long, linha As Long, ultAntes As Long
    Dim chave As String, transp As String, nf As String, v As String

    Set wsB = ThisWorkbook.Worksheets("Base_RPTG")
    Set wsR = ThisWorkbook.Worksheets("Rastreio")
    Set chaves = CreateObject("Scripting.Dictionary")
    Set cods = CreateObject("Scripting.Dictionary")
    Set nomes = CreateObject("Scripting.Dictionary")
    Set antigos = CreateObject("Scripting.Dictionary")

    colunas = Array(C_STATUS, C_OCORR, C_DTOCORR, C_PREVISAO, C_ENTREGA, C_CTE, C_CONSULTA)

    ' 1) guarda o que já foi consultado, indexado pela chave Transportadora|NF
    ultAntes = UltimaLinha(wsR, C_NF)
    For i = 2 To ultAntes
        chave = Trim$(CStr(wsR.Cells(i, C_TRANSP).Value)) & "|" & Trim$(CStr(wsR.Cells(i, C_NF).Value))
        If Not antigos.Exists(chave) Then
            ReDim guardado(0 To UBound(colunas))
            For k = 0 To UBound(colunas)
                guardado(k) = wsR.Cells(i, colunas(k)).Value
            Next k
            antigos.Add chave, guardado
        End If
    Next i

    ' 2) localiza as colunas pelo nome do cabeçalho da base
    cab = wsB.Range(wsB.Cells(1, 1), wsB.Cells(1, 40)).Value
    iStatus = Coluna(cab, "Status")
    iTransp = Coluna(cab, "Transportadora")
    iNF = Coluna(cab, "NF")
    iCod = Coluna(cab, "Codigo_STA")
    iSta = Coluna(cab, "STA")
    If iStatus = 0 Or iTransp = 0 Or iNF = 0 Then
        MsgBox "Não encontrei as colunas Status / Transportadora / NF na Base_RPTG.", vbCritical
        Exit Sub
    End If

    ' 3) varre a base em memória (rápido mesmo com 95 mil linhas)
    Application.ScreenUpdating = False
    Application.StatusBar = "Lendo a Base_RPTG..."
    ult = UltimaLinha(wsB, iTransp)
    If ult < 2 Then ult = 2
    dadosBase = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, UBound(cab, 2))).Value

    For i = 1 To UBound(dadosBase, 1)
        If InStr(1, EM_TRANSPORTE, "|" & LCase$(Trim$(CStr(dadosBase(i, iStatus)))) & "|") > 0 Then
            transp = Trim$(CStr(dadosBase(i, iTransp)))
            nf = ComoTexto(dadosBase(i, iNF))
            If Len(transp) > 0 And Len(nf) > 0 Then
                chave = transp & "|" & nf
                If Not chaves.Exists(chave) Then
                    chaves.Add chave, transp & Chr(1) & nf
                    cods.Add chave, ""
                    nomes.Add chave, ""
                End If
                If iCod > 0 Then
                    v = ComoTexto(dadosBase(i, iCod))
                    If Len(v) > 0 Then
                        If InStr(1, " " & cods(chave) & " ", " " & v & " ") = 0 Then
                            If Len(cods(chave)) = 0 Then
                                cods(chave) = v
                            Else
                                cods(chave) = cods(chave) & " / " & v
                            End If
                        End If
                    End If
                End If
                If iSta > 0 Then
                    v = Trim$(CStr(dadosBase(i, iSta)))
                    If Len(v) > 0 Then
                        If InStr(1, nomes(chave), v) = 0 Then
                            If Len(nomes(chave)) = 0 Then
                                nomes(chave) = v
                            Else
                                nomes(chave) = nomes(chave) & " / " & v
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' 4) regrava a lista
    Application.StatusBar = "Reconstruindo a aba Rastreio..."
    If ultAntes > 1 Then
        wsR.Range(wsR.Cells(2, 1), wsR.Cells(ultAntes, C_LINK)).ClearContents
    End If

    chavesOrd = ChavesOrdenadas(chaves)
    linha = 1
    If chaves.Count > 0 Then
        For i = LBound(chavesOrd) To UBound(chavesOrd)
            chave = CStr(chavesOrd(i))
            transp = Split(CStr(chaves(chave)), Chr(1))(0)
            nf = Split(CStr(chaves(chave)), Chr(1))(1)
            linha = linha + 1
            wsR.Cells(linha, C_CHAVE).Formula = "=B" & linha & "&""|""&C" & linha
            wsR.Cells(linha, C_TRANSP).Value = transp
            wsR.Cells(linha, C_NF).Value = "'" & nf
            wsR.Cells(linha, C_CODSTA).Value = cods(chave)
            wsR.Cells(linha, C_STA).Value = nomes(chave)
            wsR.Cells(linha, C_ITENS).Formula = FormulaSoma("QTD", linha)
            wsR.Cells(linha, C_VALOR).Formula = FormulaSoma("Valor_Total", linha)
            wsR.Cells(linha, C_LINK).Formula = FormulaLink(linha)
            If antigos.Exists(chave) Then
                guardado = antigos(chave)
                For k = 0 To UBound(colunas)
                    wsR.Cells(linha, colunas(k)).Value = guardado(k)
                Next k
            End If
        Next i
    End If

    AtualizarFormulaBase linha
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Aba Rastreio reconstruída: " & (linha - 1) & _
           " combinação(ões) Transportadora+NF em transporte.", vbInformation, "Rastreio"
End Sub


' ==========================================================================
'  Consulta o site público da Braspress para as linhas da aba Rastreio
' ==========================================================================
Public Sub Atualizar_Rastreio_Braspress()
    Dim wsR As Worksheet, wsP As Worksheet
    Dim cnpj As String, nfTexto As String, nf As String
    Dim linha As Long, ultima As Long, achou As Long, total As Long
    Dim html As String, partes As Variant, k As Long
    Dim st As String, ocorr As String, dtOcorr As String
    Dim prev As String, entrega As String, cte As String

    Set wsR = ThisWorkbook.Worksheets("Rastreio")
    Set wsP = ThisWorkbook.Worksheets("Parametros")

    cnpj = SomenteNumeros(CStr(wsP.Range("P4").Value))
    If Len(cnpj) < 11 Then
        MsgBox "Preencha o CNPJ do tomador do frete na aba Parametros, célula P4.", vbExclamation
        Exit Sub
    End If

    ultima = UltimaLinha(wsR, C_NF)
    Application.ScreenUpdating = False

    For linha = 2 To ultima
        If LCase$(Trim$(CStr(wsR.Cells(linha, C_TRANSP).Value))) = "braspress" Then
            nfTexto = CStr(wsR.Cells(linha, C_NF).Value)
            partes = ListaDeNotas(nfTexto)
            st = "": ocorr = "": dtOcorr = "": prev = "": entrega = "": cte = ""

            For k = LBound(partes) To UBound(partes)
                nf = CStr(partes(k))
                If Len(nf) > 0 Then
                    html = BaixarPagina(URL_BASE & "&cnpj=" & cnpj & "&numero=" & nf)
                    st = ValorTd(html, "dt-status")
                    If Len(st) > 0 Then
                        prev = ValorTd(html, "dt-previsao-entrega")
                        entrega = ValorTd(html, "dt-data-entrega")
                        cte = ValorTd(html, "dt-conhecimento")
                        UltimaEtapa html, ocorr, dtOcorr
                        Exit For
                    End If
                    Application.Wait Now + TimeSerial(0, 0, 1)
                End If
            Next k

            total = total + 1
            If Len(st) > 0 Then
                achou = achou + 1
                wsR.Cells(linha, C_STATUS).Value = st
                wsR.Cells(linha, C_OCORR).Value = ocorr
                wsR.Cells(linha, C_DTOCORR).Value = dtOcorr
                wsR.Cells(linha, C_PREVISAO).Value = prev
                wsR.Cells(linha, C_ENTREGA).Value = entrega
                wsR.Cells(linha, C_CTE).Value = cte
            Else
                wsR.Cells(linha, C_STATUS).Value = "Sem informação no site da Braspress"
            End If
            wsR.Cells(linha, C_CONSULTA).Value = Format$(Now, "dd/mm/yyyy hh:mm")
            Application.StatusBar = "Rastreio Braspress: " & total & " nota(s) consultada(s)..."
            Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next linha

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox total & " nota(s) consultada(s); " & achou & " com status retornado.", _
           vbInformation, "Rastreio Braspress"
End Sub


' ------------------------------------------------------------------ auxiliares

' última linha preenchida de uma coluna (sem depender de End(xlUp))
Private Function UltimaLinha(ByVal ws As Worksheet, ByVal Coluna As Long) As Long
    Dim ur As Range, i As Long, fim As Long
    Set ur = ws.UsedRange
    If ur Is Nothing Then
        UltimaLinha = 1
        Exit Function
    End If
    fim = ur.Row + ur.Rows.Count - 1
    For i = fim To 1 Step -1
        If Len(Trim$(CStr(ws.Cells(i, Coluna).Value))) > 0 Then
            UltimaLinha = i
            Exit Function
        End If
    Next i
    UltimaLinha = 1
End Function

Private Function FormulaSoma(ByVal campo As String, ByVal linha As Long) As String
    FormulaSoma = "=SUMIFS(tblBase[" & campo & "],tblBase[Transportadora],B" & linha & _
                  ",tblBase[NF],C" & linha & ",tblBase[Grupo_Status],""Em andamento"")"
End Function


Private Function FormulaLink(ByVal linha As Long) As String
    Dim url As String
    url = """https://www.braspress.com.br/site/w/tracking/search?cnpj=""&Parametros!$P$4&" & _
          """&documentType=NOTAFISCAL&numero=""&C" & linha
    FormulaLink = "=IF(B" & linha & "=""Braspress"",HYPERLINK(" & url & ",""Abrir rastreio""),"""")"
End Function


' mantém a coluna Status_Rastreio da base cobrindo toda a aba Rastreio
Private Sub AtualizarFormulaBase(ByVal UltimaLinha As Long)
    Dim wsB As Worksheet, lo As ListObject, lc As ListColumn, f As String
    On Error Resume Next
    Set wsB = ThisWorkbook.Worksheets("Base_RPTG")
    Set lo = wsB.ListObjects("tblBase")
    If lo Is Nothing Then Exit Sub
    Set lc = lo.ListColumns("Status_Rastreio")
    If lc Is Nothing Then Exit Sub
    If UltimaLinha < 2 Then UltimaLinha = 2
    f = "=IFERROR(INDEX(Rastreio!$H$2:$H$" & UltimaLinha & ",MATCH([@Transportadora]&""|""&[@NF]," & _
        "Rastreio!$A$2:$A$" & UltimaLinha & ",0)),"""")"
    lc.DataBodyRange.Formula = f
End Sub


Private Function Coluna(ByVal cab As Variant, ByVal nome As String) As Long
    Dim j As Long
    For j = LBound(cab, 2) To UBound(cab, 2)
        If StrComp(Trim$(CStr(cab(1, j))), nome, vbTextCompare) = 0 Then
            Coluna = j
            Exit Function
        End If
    Next j
End Function


Private Function ComoTexto(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsNumeric(v) Then
        If v = Int(v) Then
            ComoTexto = CStr(CLng(v))
            Exit Function
        End If
    End If
    ComoTexto = Trim$(CStr(v))
End Function


Private Function ChavesOrdenadas(ByVal d As Object) As Variant
    Dim arr As Variant, i As Long, j As Long, T As Variant
    If d.Count = 0 Then
        ChavesOrdenadas = Array()
        Exit Function
    End If
    arr = d.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                T = arr(i): arr(i) = arr(j): arr(j) = T
            End If
        Next j
    Next i
    ChavesOrdenadas = arr
End Function


Private Function BaixarPagina(ByVal endereco As String) As String
    Dim http As Object
    On Error GoTo falhou
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 20000, 45000
    http.Open "GET", endereco, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.Send
    If http.status = 200 Then BaixarPagina = http.responseText
    Exit Function
falhou:
    BaixarPagina = ""
End Function


Private Function ValorTd(ByVal html As String, ByVal classe As String) As String
    Dim p As Long, ini As Long, fim As Long, Txt As String
    p = InStr(1, html, "class=""" & classe & """", vbTextCompare)
    If p = 0 Then Exit Function
    ini = InStr(p, html, ">")
    fim = InStr(ini, html, "</td>")
    If ini = 0 Or fim = 0 Then Exit Function
    Txt = Trim$(Mid$(html, ini + 1, fim - ini - 1))
    If Txt = "-" Then Txt = ""
    ValorTd = Txt
End Function


Private Sub UltimaEtapa(ByVal html As String, ByRef descricao As String, ByRef data As String)
    Dim p As Long, ini As Long, fim As Long, pData As Long
    Dim d As String, dt As String
    p = 1
    Do
        p = InStr(p, html, "step-txt step-txt-up", vbTextCompare)
        If p = 0 Then Exit Do
        ini = InStr(p, html, ">")
        fim = InStr(ini, html, "</span>")
        If ini = 0 Or fim = 0 Then Exit Do
        d = LimparTexto(Mid$(html, ini + 1, fim - ini - 1))
        pData = InStr(fim, html, "step-txt step-txt-date", vbTextCompare)
        dt = ""
        If pData > 0 And pData < fim + 2000 Then
            ini = InStr(pData, html, ">")
            fim = InStr(ini, html, "</span>")
            If ini > 0 And fim > 0 Then dt = LimparTexto(Mid$(html, ini + 1, fim - ini - 1))
        End If
        If Len(d) > 0 And Len(dt) > 0 Then
            descricao = d
            data = dt
        End If
        p = p + 20
    Loop
End Sub


Private Function LimparTexto(ByVal Txt As String) As String
    Dim i As Long, ch As String, dentro As Boolean, saida As String
    For i = 1 To Len(Txt)
        ch = Mid$(Txt, i, 1)
        If ch = "<" Then
            dentro = True
        ElseIf ch = ">" Then
            dentro = False
        ElseIf Not dentro Then
            If ch = vbCr Or ch = vbLf Or ch = vbTab Then ch = " "
            saida = saida & ch
        End If
    Next i
    Do While InStr(saida, "  ") > 0
        saida = Replace(saida, "  ", " ")
    Loop
    LimparTexto = Trim$(saida)
End Function


Private Function SomenteNumeros(ByVal Txt As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(Txt)
        ch = Mid$(Txt, i, 1)
        If ch >= "0" And ch <= "9" Then SomenteNumeros = SomenteNumeros & ch
    Next i
End Function


Private Function ListaDeNotas(ByVal Txt As String) As Variant
    ' "80/83/84/85" ou "293 á 297" -> cada número citado
    Dim i As Long, ch As String, atual As String, lista As String
    For i = 1 To Len(Txt)
        ch = Mid$(Txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            atual = atual & ch
        Else
            If Len(atual) > 0 Then lista = lista & atual & "|"
            atual = ""
        End If
    Next i
    If Len(atual) > 0 Then lista = lista & atual & "|"
    If Len(lista) = 0 Then
        ListaDeNotas = Array("")
    Else
        ListaDeNotas = Split(Left$(lista, Len(lista) - 1), "|")
    End If
End Function