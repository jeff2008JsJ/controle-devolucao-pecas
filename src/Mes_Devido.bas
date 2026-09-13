Attribute VB_Name = "Mes_Devido"
Option Explicit

' ============================================================================
'  Resumo por MES DEVIDO (mes em que o prazo vence: dia 15 do 5o mes depois
'  do encerramento). Cria/atualiza duas abas:
'
'    Dados_Mes_Devido   - base auxiliar (uma linha por peca pendente)
'    Resumo_Mes_Devido  - tabela dinamica: Mes devido x Regra (cobravel/isento)
'
'  Rode Criar_Resumo_Mes_Devido sempre que quiser atualizar (depois de
'  importar o sistema de chamados, por exemplo).
'  O codigo nao usa acentos, para nao quebrar ao colar no editor VBA.
' ============================================================================

Private Const RM_BASE As String = "Base_RPTG"
Private Const RM_STAS As String = "STAs"
Private Const RM_DADOS As String = "Dados_Mes_Devido"
Private Const RM_RESUMO As String = "Resumo_Mes_Devido"
Private Const RM_DIA_LIMITE As Long = 15
Private Const RM_MESES As Long = 3
Private Const RM_UFS_NORTE As String = "AC,AP,AM,PA,RO,RR,TO"

' conversoes seguras (celulas com erro #N/D, #VALOR! nao param a macro)
Private Function RM_Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    RM_Txt = Trim$(CStr(v))
End Function

Private Function RM_Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    RM_Num = CDbl(v)
End Function

Public Sub Criar_Resumo_Mes_Devido()
    Dim n As Long

    Application.ScreenUpdating = False
    n = RM_MontarDados()
    If n = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Nenhuma peca pendente encontrada na " & RM_BASE & ".", vbExclamation
        Exit Sub
    End If
    RM_MontarDinamica n
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(RM_RESUMO).Activate
    MsgBox "Resumo por mes devido atualizado com " & n & " linha(s).", vbInformation, _
           "Devolucao RPTG"
End Sub

' ---------------------------------------------------------------------------
'  monta a aba auxiliar; devolve a quantidade de linhas gravadas
' ---------------------------------------------------------------------------
Private Function RM_MontarDados() As Long
    Dim wsB As Worksheet, wsD As Worksheet
    Dim dados As Variant, saida As Variant
    Dim cCod As Long, cNome As Long, cQtd As Long, cVal As Long, cSts As Long
    Dim cMes As Long, cAno As Long, cUF As Long, cMEI As Long, cPeca As Long
    Dim cCham As Long
    Dim i As Long, ult As Long, n As Long
    Dim ano As Long, mes As Long, prazo As Date
    Dim cod As String, uf As String, regra As String
    Dim dMei As Object, dReg As Object

    Set wsB = ThisWorkbook.Worksheets(RM_BASE)

    cCod = RM_Col(wsB, "Codigo_STA")
    cNome = RM_Col(wsB, "STA")
    cQtd = RM_Col(wsB, "QTD")
    cVal = RM_Col(wsB, "Valor_Total")
    cSts = RM_Col(wsB, "Status")
    cMes = RM_Col(wsB, "Mes")
    cAno = RM_Col(wsB, "Ano")
    cUF = RM_Col(wsB, "UF")
    cMEI = RM_Col(wsB, "MEI")
    cPeca = RM_Col(wsB, "Codigo_Peca")
    cCham = RM_Col(wsB, "Chamado")
    If cCod = 0 Or cQtd = 0 Or cSts = 0 Or cMes = 0 Or cAno = 0 Then
        MsgBox "Nao encontrei as colunas necessarias na aba " & RM_BASE & ".", vbCritical
        Exit Function
    End If

    RM_CarregarStas dMei, dReg

    ult = wsB.Cells(wsB.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Function
    dados = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, wsB.UsedRange.Columns.Count)).Value
    ReDim saida(1 To UBound(dados, 1), 1 To 12)

    For i = 1 To UBound(dados, 1)
        If RM_Norm(RM_Txt(dados(i, cSts))) = "aguardandodevolucao" Then
            ano = RM_Num(dados(i, cAno))
            mes = RM_Num(dados(i, cMes))
            cod = RM_SoNum(RM_Txt(dados(i, cCod)))
            If cod <> "" And ano > 1900 And mes >= 1 And mes <= 12 Then
                prazo = DateSerial(ano, mes + RM_MESES, RM_DIA_LIMITE)
                uf = ""
                If cUF > 0 Then uf = UCase$(RM_Txt(dados(i, cUF)))

                regra = "Cobravel"
                If dMei.Exists(cod) Then
                    regra = "Isento - MEI"
                ElseIf dReg.Exists(cod) Then
                    regra = "Isento - Norte"
                ElseIf InStr(1, "," & RM_UFS_NORTE & ",", "," & uf & ",") > 0 Then
                    regra = "Isento - Norte"
                ElseIf cMEI > 0 Then
                    If UCase$(Left$(RM_Txt(dados(i, cMEI)) & " ", 1)) = "S" Then regra = "Isento - MEI"
                End If

                n = n + 1
                saida(n, 1) = Format$(ano, "0000") & "-" & Format$(mes, "00") & " " & RM_NomeMes(mes)
                saida(n, 2) = Format$(prazo, "mm/yyyy")
                saida(n, 3) = prazo
                saida(n, 4) = IIf(Date > prazo, "Vencido", "No prazo")
                saida(n, 5) = regra
                saida(n, 6) = cod
                If cNome > 0 Then saida(n, 7) = RM_Txt(dados(i, cNome))
                saida(n, 8) = uf
                If cCham > 0 Then saida(n, 9) = RM_Txt(dados(i, cCham))
                If cPeca > 0 Then saida(n, 10) = RM_Txt(dados(i, cPeca))
                saida(n, 11) = RM_Num(dados(i, cQtd))
                If cVal > 0 Then saida(n, 12) = RM_Num(dados(i, cVal))
            End If
        End If
    Next i

    If n = 0 Then Exit Function

    Set wsD = RM_Aba(RM_DADOS)
    wsD.Cells.Clear
    wsD.Range("A1:L1").Value = Array("Mes devido", "Mes do prazo", "Prazo", "Situacao", _
                                     "Regra", "Cod STA", "STA", "UF", "Chamado", _
                                     "Cod Peca", "QTD", "Valor")
    wsD.Range(wsD.Cells(2, 1), wsD.Cells(n + 1, 12)).Value = saida
    wsD.Range(wsD.Cells(2, 3), wsD.Cells(n + 1, 3)).NumberFormat = "dd/mm/yyyy"
    wsD.Range(wsD.Cells(2, 12), wsD.Cells(n + 1, 12)).NumberFormat = "#,##0.00"
    wsD.Rows(1).Font.Bold = True
    wsD.Columns("A:L").AutoFit
    wsD.Visible = xlSheetVisible
    RM_MontarDados = n
End Function

' ---------------------------------------------------------------------------
'  tabela dinamica: linhas = mes devido, colunas = regra, valores = QTD e Valor
' ---------------------------------------------------------------------------
Private Sub RM_MontarDinamica(ByVal n As Long)
    Dim wsD As Worksheet, wsR As Worksheet
    Dim pc As PivotCache, pt As PivotTable

    Set wsD = ThisWorkbook.Worksheets(RM_DADOS)
    Set wsR = RM_Aba(RM_RESUMO)
    wsR.Cells.Clear

    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, _
             wsD.Range(wsD.Cells(1, 1), wsD.Cells(n + 1, 12)))
    Set pt = pc.CreatePivotTable(wsR.Range("A4"), "ptMesDevido")

    With pt
        .AddFields RowFields:=Array("Mes devido", "Mes do prazo"), ColumnFields:="Regra"
        .PivotFields("Situacao").Orientation = xlPageField
        With .PivotFields("QTD")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "Pecas "
            .NumberFormat = "#,##0"
        End With
        With .PivotFields("Valor")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "Valor "
            .NumberFormat = "#,##0.00"
        End With
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium2"
    End With

    ' ordena os meses pela coluna Ordem (aaaa-mm)
    On Error Resume Next
    pt.PivotFields("Mes devido").AutoSort xlAscending, "Mes devido"
    On Error GoTo 0

    wsR.Range("A1").Value = "Devolucao RPTG - pendencias por MES DEVIDO"
    wsR.Range("A2").Value = "Mes devido = mes de encerramento da peca; o prazo vence no dia " & _
                            RM_DIA_LIMITE & " do " & RM_MESES & "o mes seguinte. " & _
                            "Atualizado em " & Format$(Now, "dd/mm/yyyy hh:nn")
    wsR.Range("A1").Font.Bold = True
    wsR.Range("A1").Font.Size = 14
    wsR.Columns("A:F").AutoFit
End Sub

' ---------------------------------------------------------------------------

Private Sub RM_CarregarStas(ByRef dMei As Object, ByRef dReg As Object)
    Dim ws As Worksheet, dados As Variant, i As Long, ult As Long
    Dim cCod As Long, cMEI As Long, cReg As Long, cUF As Long, cod As String

    Set dMei = CreateObject("Scripting.Dictionary")
    Set dReg = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RM_STAS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    cCod = RM_Col(ws, "Cod STA")
    If cCod = 0 Then cCod = RM_Col(ws, "CodSTA")
    cMEI = RM_Col(ws, "MEI")
    cReg = RM_Col(ws, "Regiao")
    cUF = RM_Col(ws, "UF")
    If cCod = 0 Then Exit Sub

    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then Exit Sub
    dados = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    For i = 1 To UBound(dados, 1)
        cod = RM_SoNum(RM_Txt(dados(i, cCod)))
        If cod <> "" Then
            If cMEI > 0 Then
                If Left$(RM_Norm(RM_Txt(dados(i, cMEI))), 1) = "s" Then dMei(cod) = True
            End If
            If cReg > 0 Then
                If RM_Norm(RM_Txt(dados(i, cReg))) = "norte" Then dReg(cod) = True
            End If
            If cUF > 0 Then
                If InStr(1, "," & RM_UFS_NORTE & ",", "," & _
                   UCase$(RM_Txt(dados(i, cUF))) & ",") > 0 Then dReg(cod) = True
            End If
        End If
    Next i
End Sub

Private Function RM_Aba(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set RM_Aba = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If RM_Aba Is Nothing Then
        Set RM_Aba = ThisWorkbook.Worksheets.Add(After:= _
                     ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        RM_Aba.Name = nome
    End If
End Function

Private Function RM_Col(ws As Worksheet, ByVal titulo As String) As Long
    Dim j As Long, alvo As String
    alvo = RM_Norm(titulo)
    For j = 1 To ws.UsedRange.Columns.Count
        If RM_Norm(RM_Txt(ws.Cells(1, j).Value)) = alvo Then
            RM_Col = j
            Exit Function
        End If
    Next j
End Function

Private Function RM_SoNum(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then RM_SoNum = RM_SoNum & ch
    Next i
End Function

Private Function RM_Norm(ByVal s As String) As String
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
            Case Else:                   ch = LCase$(Mid$(s, i, 1))
        End Select
        If (ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Then r = r & ch
    Next i
    RM_Norm = r
End Function

Private Function RM_NomeMes(ByVal m As Long) As String
    Select Case m
        Case 1: RM_NomeMes = "01-Jan"
        Case 2: RM_NomeMes = "02-Fev"
        Case 3: RM_NomeMes = "03-Mar"
        Case 4: RM_NomeMes = "04-Abr"
        Case 5: RM_NomeMes = "05-Mai"
        Case 6: RM_NomeMes = "06-Jun"
        Case 7: RM_NomeMes = "07-Jul"
        Case 8: RM_NomeMes = "08-Ago"
        Case 9: RM_NomeMes = "09-Set"
        Case 10: RM_NomeMes = "10-Out"
        Case 11: RM_NomeMes = "11-Nov"
        Case 12: RM_NomeMes = "12-Dez"
    End Select
End Function