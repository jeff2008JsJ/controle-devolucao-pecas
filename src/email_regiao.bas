Attribute VB_Name = "email_regiao"
Option Explicit

' ============================================================================
'  Comunicado de devolução RPTG por REGIÃO (um e-mail por região).
'  Módulo independente: não usa nada do módulo de e-mails por STA.
' ============================================================================

Private Const RG_ABA_BASE As String = "Base_RPTG"
Private Const RG_ABA_STAS As String = "STAs"
Private Const RG_ABA_PARAM As String = "Parametros"
Private Const RG_CEL_BANNER As String = "P6"
Private Const RG_CEL_MESREF As String = "P8"
Private Const RG_CEL_LINK As String = "P9"
Private Const RG_CEL_ATENCAO As String = "P10"
Private Const RG_CEL_ANEXOS As String = "P16"      ' pasta com os PDFs a anexar
Private Const RG_EMAIL_VALIDACAO As String = "retorno.pecas@exemplo.com.br"
Private Const RG_EMAIL_ETIQUETA As String = "consultor32@exemplo.com.br"
Private Const RG_COPIA As String = "financeiro@exemplo.com.br; retorno.pecas@exemplo.com.br"
Private Const RG_ARQ_BANNER As String = "Assinatura_Empresa.png"
Private Const RG_LARGURA_BANNER As Long = 500
Private Const RG_ARQ_LOG As String = "Log_Envio_RPTG.txt"
Private Const RG_LIN_COORD_INI As Long = 12
Private Const RG_LIN_COORD_FIM As Long = 30

Private Const RG_ATENCAO_PADRAO As String = _
    "Caso algum STA esteja com devolu&ccedil;&atilde;o de {MES_REF} pendente, &eacute; necess&aacute;rio que " & _
    "encaminhe a nota at&eacute; o dia {DATA_NOTA}, pe&ccedil;as de {MES_ANT} por gentileza enviar " & _
    "at&eacute; o dia {DATA_PECAS}, sujeito a desconto se n&atilde;o for enviada no prazo informado, " & _
    "a partir de {DATA_CORTE} n&atilde;o aceitaremos notas de devolu&ccedil;&atilde;o de pe&ccedil;as de chamados " & _
    "encerrados em {MES_CORTE}, obrigado pela compreens&atilde;o e um bom {ANO}!!!"

' ---------------------------------------------------------------------------
Private Const RG_MOSTRAR_TABELA_MES As Boolean = True
Private Const RG_MOSTRAR_TABELA_STA As Boolean = False
Private Const RG_MOSTRAR_VALORES As Boolean = False
Private Const RG_MESES_TABELA As Long = 3

Private RG_SILENCIOSO As Boolean

Public Sub Revisar_Emails_Regiao()
    RG_SILENCIOSO = False
    RG_Processar False
End Sub

Public Sub Envio_Mensal_Automatico_Regiao()
    RG_SILENCIOSO = True
    RG_Log "inicio do envio mensal automatico por regiao"
    On Error Resume Next
    RG_Processar True
    If Err.Number <> 0 Then RG_Log "ERRO no envio automatico: " & Err.Description
    On Error GoTo 0
    RG_SILENCIOSO = False
End Sub

Public Sub Enviar_Emails_Regiao()
    Dim ref As Date
    ref = RG_MesReferencia()
    If MsgBox("Enviar o comunicado de " & RG_MesAno(ref) & " para todas as regi" & ChrW(245) & "es?", _
              vbYesNo + vbQuestion, "Devolu" & ChrW(231) & ChrW(227) & "o RPTG") <> vbYes Then Exit Sub
    RG_Processar True
End Sub

Public Sub Definir_Mes_Referencia()
    Dim s As String, ref As Date
    ref = RG_MesReferencia()
    s = InputBox("M" & ChrW(234) & "s de refer" & ChrW(234) & "ncia do comunicado (mm/aaaa):", "Devolu" & ChrW(231) & ChrW(227) & "o RPTG", _
                 Format$(ref, "mm/yyyy"))
    If Trim$(s) = "" Then Exit Sub
    If Not RG_LerMesAno(s, ref) Then
        MsgBox "N" & ChrW(227) & "o entendi o m" & ChrW(234) & "s. Use o formato mm/aaaa, por exemplo 05/2026.", vbExclamation
        Exit Sub
    End If
    RG_GravarParam RG_CEL_MESREF, ref
    MsgBox "M" & ChrW(234) & "s de refer" & ChrW(234) & "ncia definido: " & RG_MesAno(ref), vbInformation
End Sub

' ---------------------------------------------------------------------------

Private Sub RG_Processar(ByVal enviar As Boolean)
    Dim wsB As Worksheet, wsS As Worksheet
    Dim dados As Variant, dadosS As Variant
    Dim cCod As Long, cQtd As Long, cVal As Long, cMes As Long, cAno As Long, cGrp As Long
    Dim cSta As Long
    Dim i As Long, ult As Long, ultS As Long
    Dim regDe As Object, emailsReg As Object, pend As Object, nomeReg As Object
    Dim pendSta As Object, nomeSta As Object
    Dim cod As String, reg As String, regKey As String, chave As String
    Dim ref As Date, banner As String, temBanner As Boolean
    Dim olApp As Object, olMail As Object, olAnexo As Object
    Dim k As Variant, n As Long, falhas As Long
    Dim ac As Variant, coord As Object, cc As String
    Dim anexos As Variant, a As Variant

    Set wsB = ThisWorkbook.Worksheets(RG_ABA_BASE)
    Set wsS = ThisWorkbook.Worksheets(RG_ABA_STAS)

    cCod = RG_Coluna(wsB, "Codigo_STA")
    cQtd = RG_Coluna(wsB, "QTD")
    cVal = RG_Coluna(wsB, "Valor_Total")
    cMes = RG_Coluna(wsB, "Mes")
    cAno = RG_Coluna(wsB, "Ano")
    cGrp = RG_Coluna(wsB, "Grupo_Status")
    cSta = RG_Coluna(wsB, "STA")
    If cCod = 0 Or cQtd = 0 Or cVal = 0 Or cMes = 0 Or cAno = 0 Or cGrp = 0 Then
        If RG_SILENCIOSO Then
            RG_Log "ERRO: colunas nao encontradas na aba " & RG_ABA_BASE
        Else
            MsgBox "N" & ChrW(227) & "o encontrei as colunas necess" & ChrW(225) & "rias na aba " & RG_ABA_BASE & ".", vbCritical
        End If
        Exit Sub
    End If

    ref = RG_MesReferencia()
    Set regDe = CreateObject("Scripting.Dictionary")
    Set emailsReg = CreateObject("Scripting.Dictionary")
    Set pend = CreateObject("Scripting.Dictionary")
    Set nomeReg = CreateObject("Scripting.Dictionary")
    Set pendSta = CreateObject("Scripting.Dictionary")
    Set nomeSta = CreateObject("Scripting.Dictionary")

    ultS = wsS.Cells(wsS.Rows.Count, 6).End(xlUp).Row
    If ultS > 1 Then
        dadosS = wsS.Range(wsS.Cells(2, 1), wsS.Cells(ultS, 10)).Value
        For i = 1 To UBound(dadosS, 1)
            cod = Trim$(CStr(dadosS(i, 6)))
            reg = Trim$(CStr(dadosS(i, 4)))
            If RG_EhSim(dadosS(i, 10)) Then reg = ""
            If RG_EhNorte(reg) Then reg = ""
            If cod <> "" And reg <> "" Then
                regDe(cod) = reg
                nomeReg(LCase$(reg)) = reg
                If InStr(1, CStr(dadosS(i, 7)), "@") > 0 Then
                    regKey = LCase$(reg)
                    If Not emailsReg.Exists(regKey) Then
                        Set emailsReg(regKey) = CreateObject("Scripting.Dictionary")
                    End If
                    emailsReg(regKey)(LCase$(Trim$(CStr(dadosS(i, 7))))) = cod
                End If
            End If
        Next i
    End If

    ult = wsB.Cells(wsB.Rows.Count, cCod).End(xlUp).Row
    If ult < 2 Then
        If RG_SILENCIOSO Then
            RG_Log "ERRO: aba " & RG_ABA_BASE & " vazia"
        Else
            MsgBox "A aba " & RG_ABA_BASE & " est" & ChrW(225) & " vazia.", vbExclamation
        End If
        Exit Sub
    End If

    Application.ScreenUpdating = False
    dados = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, wsB.UsedRange.Columns.Count)).Value
    For i = 1 To UBound(dados, 1)
        Select Case Trim$(CStr(dados(i, cGrp)))
            Case "Pendente", "Em andamento"
                cod = Trim$(CStr(dados(i, cCod)))
                If regDe.Exists(cod) Then
                    chave = LCase$(CStr(regDe(cod))) & "|" & _
                            Format$(RG_Num(dados(i, cAno)), "0000") & "|" & _
                            Format$(RG_Num(dados(i, cMes)), "00")
                    If Not pend.Exists(chave) Then pend(chave) = Array(0#, 0#)
                    ac = pend(chave)
                    ac(0) = ac(0) + RG_Num(dados(i, cQtd))
                    ac(1) = ac(1) + RG_Num(dados(i, cVal))
                    pend(chave) = ac

                    chave = LCase$(CStr(regDe(cod))) & "|" & cod
                    If Not pendSta.Exists(chave) Then pendSta(chave) = Array(0#, 0#)
                    ac = pendSta(chave)
                    ac(0) = ac(0) + RG_Num(dados(i, cQtd))
                    ac(1) = ac(1) + RG_Num(dados(i, cVal))
                    pendSta(chave) = ac

                    If cSta > 0 And Not nomeSta.Exists(cod) Then
                        nomeSta(cod) = Trim$(CStr(dados(i, cSta)))
                    End If
                End If
        End Select
    Next i
    Application.ScreenUpdating = True

    banner = RG_CaminhoBanner()
    temBanner = (banner <> "")
    Set coord = RG_Consultores()
    anexos = RG_Anexos()
    If IsEmpty(anexos) Then
        If RG_SILENCIOSO Then
            RG_Log "AVISO: nenhum anexo encontrado na pasta de Parametros!" & RG_CEL_ANEXOS
        ElseIf MsgBox("Nenhum PDF encontrado na pasta de anexos (Parametros!" & _
                  RG_CEL_ANEXOS & ")." & vbCrLf & vbCrLf & _
                  "Continuar sem anexos?", vbYesNo + vbExclamation) <> vbYes Then
            Exit Sub
        End If
    End If

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        If RG_SILENCIOSO Then
            RG_Log "ERRO: Outlook classico nao disponivel nesta maquina"
        Else
            MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel abrir o Outlook nesta m" & ChrW(225) & "quina." & vbCrLf & _
                   "É preciso o Outlook cl" & ChrW(225) & "ssico instalado e logado.", vbCritical
        End If
        Exit Sub
    End If

    For Each k In RG_Ordenar(nomeReg.Keys)
        reg = CStr(nomeReg(k))
        If Not emailsReg.Exists(CStr(k)) Then GoTo proxima
        If emailsReg(CStr(k)).Count = 0 Then GoTo proxima

        On Error Resume Next
        Set olMail = olApp.CreateItem(0)
        With olMail
            cc = RG_COPIA
            If coord.Exists(CStr(k)) Then cc = cc & "; " & CStr(coord(CStr(k)))
            .To = Join(emailsReg(CStr(k)).Keys, "; ")
            .cc = cc
            .Importance = 2
            .Subject = "N" & ChrW(195) & "O RESPONDA - DEVOLU" & ChrW(199) & ChrW(195) & "O RPTG - CHAMADOS " & _
                       UCase$(RG_NomeMes(Month(ref))) & "/" & Year(ref) & " " & UCase$(reg)
            If temBanner Then
                Set olAnexo = .Attachments.Add(banner, 1, 0)
                olAnexo.PropertyAccessor.SetProperty _
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "bannerrptg"
                Set olAnexo = Nothing
            End If
            If Not IsEmpty(anexos) Then
                For Each a In anexos
                    .Attachments.Add CStr(a)
                Next a
            End If
            .HTMLBody = RG_Html(reg, CStr(k), ref, pend, temBanner, pendSta, nomeSta)
            If enviar Then .Send Else .Display
        End With
        If Err.Number <> 0 Then
            falhas = falhas + 1
            RG_Log "ERRO na regiao " & reg & ": " & Err.Description
            Err.Clear
        Else
            n = n + 1
        End If
        On Error GoTo 0
        Set olMail = Nothing
        If enviar Then Application.Wait Now + TimeSerial(0, 0, 1)
proxima:
    Next k

    If enviar And n > 0 Then
        RG_GravarParam RG_CEL_MESREF, DateSerial(Year(ref), Month(ref) + 1, 1)
        RG_Log "comunicado de " & RG_MesAno(ref) & " enviado para " & n & " regiao(oes)"
    End If

    If RG_SILENCIOSO Then
        RG_Log "fim do envio automatico - enviados: " & n & " / falhas: " & falhas
        Exit Sub
    End If

    MsgBox IIf(enviar, "E-mails enviados: ", "E-mails abertos para revis" & ChrW(227) & "o: ") & n & _
           IIf(falhas > 0, vbCrLf & "Falhas: " & falhas, "") & vbCrLf & vbCrLf & _
           "M" & ChrW(234) & "s de refer" & ChrW(234) & "ncia usado: " & RG_MesAno(ref) & _
           IIf(enviar And n > 0, vbCrLf & "Pr" & ChrW(243) & "ximo envio: " & _
               RG_MesAno(DateSerial(Year(ref), Month(ref) + 1, 1)), ""), vbInformation
End Sub

Private Function RG_Cabec_Valor() As String
    If RG_MOSTRAR_VALORES Then
        RG_Cabec_Valor = "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">Valor</td>"
    End If
End Function

Private Function RG_Celula_Valor(ByVal v As Double) As String
    If RG_MOSTRAR_VALORES Then
        RG_Celula_Valor = "<td style=""border:1px solid #d0d0d0;padding:3px 8px;" & _
                          "text-align:right;"">R$ " & Format$(v, "#,##0.00") & "</td>"
    End If
End Function

' ---------------------------------------------------------------------------
'  Converte caracteres acentuados para Entidades HTML Universais
' ---------------------------------------------------------------------------
Private Function RG_Entidades(ByVal s As String) As String
    s = Replace(s, "á", "&aacute;")
    s = Replace(s, "Á", "&Aacute;")
    s = Replace(s, "à", "&agrave;")
    s = Replace(s, "À", "&Agrave;")
    s = Replace(s, "â", "&acirc;")
    s = Replace(s, "Â", "&Acirc;")
    s = Replace(s, "ã", "&atilde;")
    s = Replace(s, "Ã", "&Atilde;")
    s = Replace(s, "é", "&eacute;")
    s = Replace(s, "É", "&Eacute;")
    s = Replace(s, "ê", "&ecirc;")
    s = Replace(s, "Ê", "&Ecirc;")
    s = Replace(s, "í", "&iacute;")
    s = Replace(s, "Í", "&Iacute;")
    s = Replace(s, "ó", "&oacute;")
    s = Replace(s, "Ó", "&Oacute;")
    s = Replace(s, "ô", "&ocirc;")
    s = Replace(s, "Ô", "&Ocirc;")
    s = Replace(s, "õ", "&otilde;")
    s = Replace(s, "Õ", "&Otilde;")
    s = Replace(s, "ú", "&uacute;")
    s = Replace(s, "Ú", "&Uacute;")
    s = Replace(s, "ç", "&ccedil;")
    s = Replace(s, "Ç", "&Ccedil;")
    RG_Entidades = s
End Function

' ---------------------------------------------------------------------------
'  Corpo do e-mail em HTML (100% blindado com Entidades HTML)
' ---------------------------------------------------------------------------
Private Function RG_Html(ByVal reg As String, ByVal regKey As String, ByVal ref As Date, _
                         pend As Object, ByVal temBanner As Boolean, _
                         pendSta As Object, nomeSta As Object) As String
    Dim h As String, saudacao As String, link As String, atencao As String
    Dim resumo As String, k As Variant, ac As Variant
    Dim totQtd As Double, totVal As Double
    Dim porSta As String, cod As String, nm As String
    Dim qtdMeses As Long, pular As Long
    Dim qtdSta As Double, valSta As Double
    Dim marca As String, ini As String, fim As String

    If Hour(Now) < 12 Then saudacao = "Bom Dia" Else saudacao = "Boa Tarde"
    marca = "background-color:#ffff00;"
    ini = "<p style=""margin:0 0 10px 0;"">"
    fim = "</p>"

    link = RG_Param(RG_CEL_LINK)
    If link <> "" Then
        link = "<a href=""" & link & """>TBL_COD_RPTG - Planilhas Google</a>"
    Else
        link = "TBL_COD_RPTG - Planilhas Google"
    End If

    h = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f1f1f;line-height:1.4;"">"
    h = h & "<p style=""text-align:center;font-weight:bold;margin:0 0 14px 0;"">" & _
        saudacao & " ! STA's,</p>"
    h = h & ini & "Por gentileza fazer devolu&ccedil;&atilde;o de pe&ccedil;as RPTG referente ao per&iacute;odo de " & _
        "encerramento de <b>" & RG_Entidades(RG_NomeMes(Month(ref))) & " de " & Year(ref) & "</b>." & fim
    h = h & ini & "Por favor seguir instru&ccedil;&otilde;es:" & fim
    h = h & "<ul style=""margin:0 0 10px 22px;padding:0;"">"
    h = h & "<li>Verificar devolu&ccedil;&otilde;es pendentes no sistema de chamados, caso n&atilde;o tenha realizado, " & _
        "emitir espelho (Pr&eacute;-Danfe) das notas e encaminhar, <b>(Emitir na sefaz apenas " & _
        "ap&oacute;s valida&ccedil;&atilde;o do espelho da NF)</b>;</li>"
    h = h & "<li><span style=""" & marca & """>O pr&eacute;-danfe deve conter os chamados " & _
        "encerrados do in&iacute;cio at&eacute; o fim de cada m&ecirc;s, n&atilde;o ser&aacute; aceita a devolu&ccedil;&atilde;o " & _
        "parcial/incompleta;</span></li>"
    h = h & "<li>Certificar que os c&oacute;digos gen&eacute;ricos, descri&ccedil;&atilde;o e valores de devolu&ccedil;&atilde;o da NF " & _
        "estejam corretos, essas informa&ccedil;&otilde;es podem ser consultadas utilizando planilha no " & _
        "link: " & link & ";</li>"
    h = h & "<li>Pe&ccedil;as de Misturadores/Chopeiras, e Geladeiras devem ser emitidas em notas " & _
        "separadas;</li>"
    h = h & "<li><span style=""" & marca & """>Todos os compressores devem voltar com a " & _
        "parte el&eacute;trica, rel&eacute; e protetor.</span></li>"
    h = h & "</ul>"

    h = h & ini & "Para valida&ccedil;&atilde;o de NFs, enviar e-mail diretamente ao: " & _
        "<a href=""mailto:" & RG_EMAIL_VALIDACAO & """>" & RG_EMAIL_VALIDACAO & "</a> " & _
        "<span style=""color:#c00000;"">(Mencionar no e-mail o m&ecirc;s de encerramento dos " & _
        "chamados das pe&ccedil;as na NF)</span><br>" & _
        "Para solicita&ccedil;&atilde;o de etiqueta RPTG, enviar e-mail para " & _
        "<a href=""mailto:" & RG_EMAIL_ETIQUETA & """>" & RG_EMAIL_ETIQUETA & "</a>" & fim

    h = h & ini & "<span style=""" & marca & """><i>OBS: caso j&aacute; tenha devolu&ccedil;&atilde;o deste m&ecirc;s " & _
        "em andamento, por gentileza desconsiderar este e-mail.</i></span>" & fim

    h = h & "<p style=""text-align:center;margin:14px 0 6px 0;"">" & _
        "<span style=""background-color:#c00000;color:#ffffff;font-weight:bold;" & _
        "padding:2px 10px;"">Aten&ccedil;&atilde;o!</span></p>"

    atencao = RG_Param(RG_CEL_ATENCAO)
    If atencao = "" Then
        atencao = RG_ATENCAO_PADRAO
        RG_GravarParam RG_CEL_ATENCAO, atencao
    End If
    atencao = RG_Entidades(atencao)
    atencao = Replace(atencao, "{MES_REF}", RG_Entidades(RG_MesAnoBarra(ref)))
    atencao = Replace(atencao, "{MES_ANT}", RG_Entidades(RG_NomeMes(Month(RG_SomaMes(ref, -1)))))
    atencao = Replace(atencao, "{MES_CORTE}", RG_Entidades(RG_MesAno(RG_SomaMes(ref, -2))))
    atencao = Replace(atencao, "{DATA_NOTA}", _
        Format$(DateSerial(Year(Date), Month(Date) + 1, 0), "dd/mm/yyyy"))
    atencao = Replace(atencao, "{DATA_PECAS}", _
        Format$(DateSerial(Year(Date), Month(Date) + 1, 15), "dd/mm/yyyy"))
    atencao = Replace(atencao, "{DATA_CORTE}", _
        Format$(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yyyy"))
    atencao = Replace(atencao, "{ANO}", CStr(Year(Date)))
    h = h & "<p style=""margin:0 0 12px 0;font-style:italic;"">" & atencao & fim

    ' --- resumo das pendências da região
    resumo = ""
    If Not RG_MOSTRAR_TABELA_MES Then GoTo semMes

    For Each k In pend.Keys
        If Left$(CStr(k), Len(regKey) + 1) = regKey & "|" Then
            ac = pend(k)
            If ac(0) > 0 Then qtdMeses = qtdMeses + 1
        End If
    Next k
    If RG_MESES_TABELA > 0 And qtdMeses > RG_MESES_TABELA Then
        pular = qtdMeses - RG_MESES_TABELA
    End If

    For Each k In RG_Ordenar(pend.Keys)
        If Left$(CStr(k), Len(regKey) + 1) = regKey & "|" Then
            ac = pend(k)
            If ac(0) > 0 And pular > 0 Then
                pular = pular - 1
                ac = Array(0#, 0#)
            End If
            If ac(0) > 0 Then
                resumo = resumo & "<tr><td style=""border:1px solid #d0d0d0;padding:3px 8px;"">" & _
                    Mid$(CStr(k), Len(regKey) + 7, 2) & "/" & Mid$(CStr(k), Len(regKey) + 2, 4) & _
                    "</td><td style=""border:1px solid #d0d0d0;padding:3px 8px;text-align:right;"">" & _
                    Format$(ac(0), "#,##0") & "</td>" & _
                    RG_Celula_Valor(ac(1)) & "</tr>"
                totQtd = totQtd + ac(0)
                totVal = totVal + ac(1)
            End If
        End If
    Next k
    If resumo <> "" Then
        h = h & ini & "<b>Pend&ecirc;ncias em aberto da regi&atilde;o " & RG_Entidades(reg) & _
            IIf(RG_MESES_TABELA > 0, " (&uacute;ltimos " & RG_MESES_TABELA & " meses)", "") & _
            ":</b>" & fim
        h = h & "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;" & _
            "font-size:10pt;margin:0 0 12px 0;"">" & _
            "<tr style=""background-color:#f2f2f2;font-weight:bold;"">" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">M&ecirc;s/Ano</td>" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">Pe&ccedil;as</td>" & _
            RG_Cabec_Valor() & "</tr>" & _
            resumo & "<tr style=""font-weight:bold;"">" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">Total</td>" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;text-align:right;"">" & _
            Format$(totQtd, "#,##0") & "</td>" & _
            RG_Celula_Valor(totVal) & "</tr></table>"
    End If
semMes:

    ' --- pendências em aberto de cada STA
    porSta = ""
    If Not RG_MOSTRAR_TABELA_STA Then GoTo semSta
    For Each k In RG_Ordenar(pendSta.Keys)
        If Left$(CStr(k), Len(regKey) + 1) = regKey & "|" Then
            ac = pendSta(k)
            If ac(0) > 0 Then
                cod = Mid$(CStr(k), Len(regKey) + 2)
                nm = ""
                If nomeSta.Exists(cod) Then nm = CStr(nomeSta(cod))
                qtdSta = qtdSta + ac(0)
                valSta = valSta + ac(1)
                porSta = porSta & "<tr><td style=""border:1px solid #d0d0d0;padding:3px 8px;"">" & _
                    cod & "</td><td style=""border:1px solid #d0d0d0;padding:3px 8px;"">" & RG_Entidades(nm) & "</td>" & _
                    "<td style=""border:1px solid #d0d0d0;padding:3px 8px;text-align:right;"">" & _
                    Format$(ac(0), "#,##0") & "</td>" & _
                    RG_Celula_Valor(ac(1)) & "</tr>"
            End If
        End If
    Next k
    If porSta <> "" Then
        h = h & ini & "<b>Pend&ecirc;ncias por STA (" & RG_Entidades(reg) & "):</b> " & _
            "<span style=""font-size:9pt;color:#606060;"">localize o seu c&oacute;digo na lista abaixo.</span>" & fim
        h = h & "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;" & _
            "font-size:10pt;margin:0 0 12px 0;"">" & _
            "<tr style=""background-color:#f2f2f2;font-weight:bold;"">" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">C&oacute;d STA</td>" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">STA</td>" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"">Pe&ccedil;as</td>" & _
            RG_Cabec_Valor() & "</tr>" & _
            porSta & "<tr style=""font-weight:bold;"">" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;"" colspan=""2"">Total</td>" & _
            "<td style=""border:1px solid #d0d0d0;padding:3px 8px;text-align:right;"">" & _
            Format$(qtdSta, "#,##0") & "</td>" & _
            RG_Celula_Valor(valSta) & "</tr></table>"
    End If
semSta:

    h = h & "<p style=""margin:14px 0 0 0;"">" & _
        "<b style=""color:#1f4e79;"">Atenciosamente,</b><br>" & _
        "<b style=""color:#1f4e79;"">Nome do Analista</b><br>" & _
        "<span style=""color:#808080;"">Analista de Qualidade - RPTG</span><br>" & _
        "<span style=""color:#808080;"">S&atilde;o Paulo - Brasil</span><br>" & _
        "E-mail: <a href=""mailto:analista.rptg@exemplo.com.br"">analista.rptg@exemplo.com.br</a>" & _
        "</p>"

    If temBanner Then
        h = h & "<p style=""margin:10px 0 0 0;""><img src=""cid:bannerrptg"" width=""" & _
            RG_LARGURA_BANNER & """ style=""border:0;display:block;""></p>"
    End If

    RG_Html = "<!DOCTYPE html><html><head><meta http-equiv=""Content-Type"" " & _
              "content=""text/html; charset=utf-8""></head><body>" & _
              h & "</div></body></html>"
End Function

' ---------------------------------------------------------------------------
'  Apoio
' ---------------------------------------------------------------------------

Private Function RG_EhSim(ByVal v As Variant) As Boolean
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    RG_EhSim = (s = "sim" Or s = "s" Or s = "true" Or s = "verdadeiro" Or s = "1")
End Function

Private Function RG_EhNorte(ByVal reg As String) As Boolean
    RG_EhNorte = (InStr(1, LCase$(reg), "norte") > 0 And InStr(1, LCase$(reg), "nordeste") = 0)
End Function

Private Function RG_Num(ByVal v As Variant) As Double
    If IsNumeric(v) Then RG_Num = CDbl(v)
End Function

Private Function RG_NomeMes(ByVal m As Long) As String
    Dim nomes As Variant
    nomes = Array("Janeiro", "Fevereiro", "Mar" & ChrW(231) & "o", "Abril", "Maio", "Junho", _
                  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
    If m >= 1 And m <= 12 Then RG_NomeMes = nomes(m - 1)
End Function

Private Function RG_MesAno(ByVal d As Date) As String
    RG_MesAno = RG_NomeMes(Month(d)) & "/" & Year(d)
End Function

Private Function RG_MesAnoBarra(ByVal d As Date) As String
    RG_MesAnoBarra = RG_NomeMes(Month(d)) & "/" & Year(d)
End Function

Private Function RG_SomaMes(ByVal d As Date, ByVal n As Long) As Date
    RG_SomaMes = DateSerial(Year(d), Month(d) + n, 1)
End Function

Private Function RG_LerMesAno(ByVal s As String, ByRef ref As Date) As Boolean
    Dim p() As String
    s = Replace(Replace(Trim$(s), "-", "/"), ".", "/")
    p = Split(s, "/")
    If UBound(p) < 1 Then Exit Function
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(UBound(p))) Then Exit Function
    On Error Resume Next
    ref = DateSerial(CLng(p(UBound(p))), CLng(p(0)), 1)
    RG_LerMesAno = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function RG_MesReferencia() As Date
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(RG_ABA_PARAM).Range(RG_CEL_MESREF).Value
    On Error GoTo 0
    If IsDate(v) Then
        RG_MesReferencia = DateSerial(Year(v), Month(v), 1)
    Else
        RG_MesReferencia = DateSerial(Year(Date), Month(Date) - 1, 1)
        RG_GravarParam RG_CEL_MESREF, RG_MesReferencia
    End If
End Function

Private Function RG_Anexos() As Variant
    Dim pasta As String, arq As String, lista As Collection, i As Long
    Dim res() As String
    pasta = RG_Param(RG_CEL_ANEXOS)
    If pasta = "" Then Exit Function
    If Right$(pasta, 1) <> Application.PathSeparator Then _
        pasta = pasta & Application.PathSeparator
    If Dir(pasta, vbDirectory) = "" Then Exit Function
    Set lista = New Collection
    arq = Dir(pasta & "*.*")
    Do While arq <> ""
        If Left$(arq, 1) <> "~" Then lista.Add pasta & arq
        arq = Dir
    Loop
    If lista.Count = 0 Then Exit Function
    ReDim res(0 To lista.Count - 1)
    For i = 1 To lista.Count
        res(i - 1) = lista(i)
    Next i
    RG_Anexos = res
End Function

Public Sub Escolher_Pasta_Anexos()
    Dim fd As Object, pasta As String
    Set fd = Application.FileDialog(4)
    fd.Title = "Selecione a pasta com os PDFs do comunicado"
    If fd.Show <> -1 Then Exit Sub
    pasta = fd.SelectedItems(1)
    RG_GravarParam RG_CEL_ANEXOS, pasta
    MsgBox "Pasta de anexos definida:" & vbCrLf & pasta & vbCrLf & vbCrLf & _
           "Arquivos encontrados: " & RG_QtdAnexos(), vbInformation
End Sub

Private Function RG_QtdAnexos() As Long
    Dim v As Variant
    v = RG_Anexos()
    If Not IsEmpty(v) Then RG_QtdAnexos = UBound(v) - LBound(v) + 1
End Function

Private Function RG_Consultores() As Object
    Dim ws As Worksheet, i As Long, reg As String, lista As String
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RG_ABA_PARAM)
    On Error GoTo 0
    If Not ws Is Nothing Then
        For i = RG_LIN_COORD_INI To RG_LIN_COORD_FIM
            reg = Trim$(CStr(ws.Cells(i, 15).Value))
            lista = Trim$(CStr(ws.Cells(i, 16).Value))
            If reg <> "" And InStr(1, lista, "@") > 0 Then
                dic(LCase$(reg)) = lista
            End If
        Next i
    End If
    Set RG_Consultores = dic
End Function

Private Function RG_Param(ByVal cel As String) As String
    On Error Resume Next
    RG_Param = Trim$(CStr(ThisWorkbook.Worksheets(RG_ABA_PARAM).Range(cel).Value))
    On Error GoTo 0
End Function

Private Sub RG_GravarParam(ByVal cel As String, ByVal valor As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RG_ABA_PARAM)
    ws.Range(cel).Value = valor
    Select Case cel
        Case RG_CEL_MESREF
            ws.Range(cel).NumberFormat = "mm/yyyy"
            If Trim$(CStr(ws.Range("O8").Value)) = "" Then _
                ws.Range("O8").Value = "M" & ChrW(234) & "s de refer" & ChrW(234) & "ncia do comunicado"
        Case RG_CEL_LINK
            If Trim$(CStr(ws.Range("O9").Value)) = "" Then _
                ws.Range("O9").Value = "Link TBL_COD_RPTG"
        Case RG_CEL_ATENCAO
            If Trim$(CStr(ws.Range("O10").Value)) = "" Then _
                ws.Range("O10").Value = "Texto do bloco Aten" & ChrW(231) & ChrW(227) & "o!"
        Case RG_CEL_ANEXOS
            If Trim$(CStr(ws.Range("O16").Value)) = "" Then _
                ws.Range("O16").Value = "Pasta dos PDFs anexados"
    End Select
    On Error GoTo 0
End Sub

Private Function RG_Coluna(ws As Worksheet, ByVal titulo As String) As Long
    Dim c As Long, ultC As Long
    ultC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultC
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), titulo, vbTextCompare) = 0 Then
            RG_Coluna = c
            Exit Function
        End If
    Next c
End Function

Private Function RG_Ordenar(ByVal chaves As Variant) As Variant
    Dim i As Long, j As Long, T As Variant
    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If CStr(chaves(j)) < CStr(chaves(i)) Then
                T = chaves(i): chaves(i) = chaves(j): chaves(j) = T
            End If
        Next j
    Next i
    RG_Ordenar = chaves
End Function

Private Function RG_PastaLocal() As String
    Dim p As String, resto As String, i As Long
    Dim base As Variant
    p = ThisWorkbook.Path
    If p = "" Then Exit Function
    If LCase$(Left$(p, 4)) <> "http" Then
        RG_PastaLocal = p
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
                RG_PastaLocal = CStr(base) & resto
                Exit Function
            End If
        End If
    Next base
End Function

Private Function RG_CaminhoBanner() As String
    Dim caminho As String, pasta As String
    caminho = RG_Param(RG_CEL_BANNER)
    If caminho <> "" Then
        If Dir(caminho) <> "" Then
            RG_CaminhoBanner = caminho
            Exit Function
        End If
    End If
    pasta = RG_PastaLocal()
    If pasta <> "" Then
        caminho = pasta & Application.PathSeparator & RG_ARQ_BANNER
        If Dir(caminho) <> "" Then RG_CaminhoBanner = caminho
    End If
End Function

Private Sub RG_Log(ByVal texto As String)
    Dim pasta As String, arq As String, fnum As Integer
    On Error Resume Next
    pasta = RG_PastaLocal()
    If pasta = "" Then pasta = Environ$("TEMP")
    arq = pasta & Application.PathSeparator & RG_ARQ_LOG
    fnum = FreeFile
    Open arq For Append As #fnum
    Print #fnum, Format$(Now, "dd/mm/yyyy hh:nn") & " - [regiao] " & texto
    Close #fnum
    On Error GoTo 0
End Sub