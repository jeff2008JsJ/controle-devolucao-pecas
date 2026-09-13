Attribute VB_Name = "Cobranca_Semanal"
Option Explicit

' ============================================================================
'  Cobranca automatica SEM o Agendador de Tarefas do Windows.
'  A propria planilha marca a hora (Application.OnTime) e dispara o envio
'  chamando Envio_Semanal_Automatico. Funciona enquanto a planilha estiver
'  aberta; se ela estiver fechada na hora marcada, a cobranca acontece
'  alguns minutos depois de abrir (se ja tiver passado o intervalo).
'
'  Rotinas para usar:
'    Configurar_Agendamento      - escolhe o dia da semana, a hora e o intervalo
'    Ativar_Cobranca_Semanal     - liga o agendamento (ou reprograma)
'    Desativar_Cobranca_Semanal  - desliga
'    Situacao_Cobranca_Semanal   - mostra o proximo horario marcado
'
'  O dia/hora/intervalo escolhidos ficam gravados na aba Parametros
'  (P24, P25 e P26), entao nao e' preciso mexer no codigo.
' ============================================================================

Private Const AG_DIA_PADRAO As Long = 2        ' 1=domingo, 2=segunda ... 7=sabado
Private Const AG_HORA_PADRAO As String = "09:00"
Private Const AG_DIAS_PADRAO As Long = 7       ' periodicidade, em dias
Private Const AG_ESPERA_MIN As Long = 3        ' minutos depois de abrir, no atraso
Private Const AG_ABA_PARAM As String = "Parametros"
Private Const AG_CEL_ULT As String = "P23"     ' ultima execucao automatica
Private Const AG_CEL_ROT As String = "O23"     ' rotulo da celula acima
Private Const AG_CEL_DIA As String = "P24"     ' dia da semana escolhido
Private Const AG_CEL_HORA As String = "P25"    ' hora escolhida
Private Const AG_CEL_DIAS As String = "P26"    ' intervalo em dias
Private Const AG_MACRO As String = "Rodar_Cobranca_Semanal"

Private ag_proxima As Double                   ' horario marcado (0 = nada marcado)

' ---------------------------------------------------------------------------
'  Liga o agendamento. Rode uma vez; para ligar sozinho quando a planilha
'  abrir, coloque no modulo EstaPastaDeTrabalho:
'
'     Private Sub Workbook_Open()
'         Ativar_Cobranca_Semanal
'     End Sub
' ---------------------------------------------------------------------------
Public Sub Ativar_Cobranca_Semanal()
    AG_Programar True
End Sub

Public Sub Desativar_Cobranca_Semanal()
    AG_Cancelar
    MsgBox "Cobranca automatica desligada.", vbInformation, "Devolucao RPTG"
End Sub

' ---------------------------------------------------------------------------
'  Escolhe dia da semana, hora e intervalo; grava na aba Parametros e reagenda
' ---------------------------------------------------------------------------
Public Sub Configurar_Agendamento()
    Dim r As String, dia As Long, hora As String, dias As Long

    r = InputBox("Qual dia da semana enviar?" & vbCrLf & vbCrLf & _
                 "1 = domingo" & vbCrLf & "2 = segunda" & vbCrLf & _
                 "3 = terca" & vbCrLf & "4 = quarta" & vbCrLf & _
                 "5 = quinta" & vbCrLf & "6 = sexta" & vbCrLf & "7 = sabado", _
                 "Agendamento da cobranca", CStr(AG_DiaSemana()))
    If Trim$(r) = "" Then Exit Sub
    dia = Val(r)
    If dia < 1 Or dia > 7 Then
        MsgBox "Informe um numero de 1 a 7.", vbExclamation
        Exit Sub
    End If

    r = InputBox("Qual horario? (formato hh:mm, ex.: 09:00)", _
                 "Agendamento da cobranca", AG_HoraTxt())
    If Trim$(r) = "" Then Exit Sub
    hora = Trim$(r)
    If Not IsDate(hora) Then
        MsgBox "Horario invalido. Use o formato hh:mm, ex.: 15:30.", vbExclamation
        Exit Sub
    End If
    hora = Format$(TimeValue(hora), "hh:nn")

    r = InputBox("De quantos em quantos dias?" & vbCrLf & vbCrLf & _
                 "7 = toda semana" & vbCrLf & "15 = a cada 15 dias" & vbCrLf & _
                 "30 = uma vez por mes", _
                 "Agendamento da cobranca", CStr(AG_Dias()))
    If Trim$(r) = "" Then Exit Sub
    dias = Val(r)
    If dias < 1 Then
        MsgBox "Informe um numero de dias maior que zero.", vbExclamation
        Exit Sub
    End If

    AG_Gravar AG_CEL_DIA, dia, "Cobranca automatica - dia da semana (1=dom ... 7=sab)", "0"
    AG_Gravar AG_CEL_HORA, hora, "Cobranca automatica - horario", "@"
    AG_Gravar AG_CEL_DIAS, dias, "Cobranca automatica - intervalo em dias", "0"

    AG_Programar True
End Sub

Public Sub Situacao_Cobranca_Semanal()
    Dim ult As String

    If IsDate(AG_Ultima()) Then
        ult = Format$(AG_Ultima(), "dd/mm/yyyy hh:nn")
    Else
        ult = "nunca"
    End If

    If ag_proxima = 0 Then
        MsgBox "A cobranca automatica esta DESLIGADA." & vbCrLf & _
               "Dia: " & AG_NomeDia(AG_DiaSemana()) & "   Hora: " & AG_HoraTxt() & _
               "   Intervalo: " & AG_Dias() & " dia(s)" & vbCrLf & _
               "Ultimo envio automatico: " & ult & vbCrLf & vbCrLf & _
               "Rode Ativar_Cobranca_Semanal para ligar.", vbInformation, "Devolucao RPTG"
    Else
        MsgBox "Cobranca automatica LIGADA." & vbCrLf & _
               "Dia: " & AG_NomeDia(AG_DiaSemana()) & "   Hora: " & AG_HoraTxt() & _
               "   Intervalo: " & AG_Dias() & " dia(s)" & vbCrLf & _
               "Proximo envio: " & Format$(ag_proxima, "dd/mm/yyyy hh:nn") & vbCrLf & _
               "Ultimo envio automatico: " & ult & vbCrLf & vbCrLf & _
               "A planilha precisa estar aberta na hora marcada." & vbCrLf & _
               "Para mudar dia/hora, rode Configurar_Agendamento.", vbInformation, "Devolucao RPTG"
    End If
End Sub

' rotina disparada pelo OnTime: envia os dois e-mails e ja marca o proximo
'   Envio_Semanal_Automatico = cobranca de coleta (STAs "Em andamento")
'   Envio_Semanal_Devolucao  = cobranca de devolucao ("Aguardando devolucao")
' Application.Run evita erro de compilacao se um dos modulos nao estiver na planilha.
Public Sub Rodar_Cobranca_Semanal()
    ag_proxima = 0

    On Error Resume Next
    Application.Run "Envio_Semanal_Automatico"
    Err.Clear
    Application.Run "Envio_Semanal_Devolucao"
    Err.Clear
    AG_GravarUltima Now
    On Error GoTo 0

    AG_Programar False
End Sub

' ---------------------------------------------------------------------------

Private Sub AG_Programar(ByVal avisar As Boolean)
    Dim quando As Double

    AG_Cancelar
    quando = AG_ProximoHorario()
    On Error Resume Next
    Application.OnTime quando, AG_MACRO
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        If avisar Then MsgBox "Nao foi possivel marcar o envio automatico.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    ag_proxima = quando

    If avisar Then
        MsgBox "Cobranca automatica ligada." & vbCrLf & _
               "Dia: " & AG_NomeDia(AG_DiaSemana()) & "   Hora: " & AG_HoraTxt() & _
               "   Intervalo: " & AG_Dias() & " dia(s)" & vbCrLf & _
               "Proximo envio: " & Format$(quando, "dd/mm/yyyy hh:nn") & vbCrLf & vbCrLf & _
               "Deixe a planilha aberta nesse horario." & vbCrLf & _
               "Para mudar dia/hora, rode Configurar_Agendamento.", vbInformation, "Devolucao RPTG"
    End If
End Sub

Private Sub AG_Cancelar()
    If ag_proxima = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime ag_proxima, AG_MACRO, , False
    On Error GoTo 0
    ag_proxima = 0
End Sub

' proximo horario: o dia da semana escolhido, na hora escolhida.
' Se ja passou o intervalo desde o ultimo envio, roda logo depois de abrir.
Private Function AG_ProximoHorario() As Double
    Dim alvo As Date, hora As Date, dias As Long, ult As Variant

    hora = TimeValue(AG_HoraTxt())
    ult = AG_Ultima()

    If IsDate(ult) Then
        If Now - CDate(ult) >= AG_Dias() Then
            AG_ProximoHorario = CDbl(Now + TimeSerial(0, AG_ESPERA_MIN, 0))
            Exit Function
        End If
    End If

    dias = AG_DiaSemana() - Weekday(Date)
    If dias < 0 Then dias = dias + 7
    alvo = Date + dias + hora
    If alvo <= Now Then alvo = alvo + 7
    AG_ProximoHorario = CDbl(alvo)
End Function

' ---------------------------------------------------------------------------
' configuracao gravada na aba Parametros (se estiver vazia, usa o padrao)

Private Function AG_DiaSemana() As Long
    Dim v As Variant

    v = AG_Ler(AG_CEL_DIA)
    AG_DiaSemana = AG_DIA_PADRAO
    If IsNumeric(v) Then
        If Val(v) >= 1 And Val(v) <= 7 Then AG_DiaSemana = CLng(Val(v))
    End If
End Function

Private Function AG_HoraTxt() As String
    Dim v As Variant

    v = AG_Ler(AG_CEL_HORA)
    AG_HoraTxt = AG_HORA_PADRAO
    On Error Resume Next
    If IsDate(v) Then AG_HoraTxt = Format$(TimeValue(CStr(v)), "hh:nn")
    On Error GoTo 0
End Function

Private Function AG_Dias() As Long
    Dim v As Variant

    v = AG_Ler(AG_CEL_DIAS)
    AG_Dias = AG_DIAS_PADRAO
    If IsNumeric(v) Then
        If Val(v) >= 1 Then AG_Dias = CLng(Val(v))
    End If
End Function

Private Function AG_Ler(ByVal cel As String) As Variant
    On Error Resume Next
    AG_Ler = ThisWorkbook.Worksheets(AG_ABA_PARAM).Range(cel).Value
    On Error GoTo 0
End Function

Private Sub AG_Gravar(ByVal cel As String, ByVal valor As Variant, _
                      ByVal rotulo As String, ByVal formato As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AG_ABA_PARAM)
    If ws Is Nothing Then Exit Sub
    ws.Range("O" & Mid$(cel, 2)).Value = rotulo
    ws.Range(cel).NumberFormat = formato
    ws.Range(cel).Value = valor
    On Error GoTo 0
End Sub

Private Function AG_NomeDia(ByVal d As Long) As String
    Select Case d
        Case 1: AG_NomeDia = "domingo"
        Case 2: AG_NomeDia = "segunda"
        Case 3: AG_NomeDia = "terca"
        Case 4: AG_NomeDia = "quarta"
        Case 5: AG_NomeDia = "quinta"
        Case 6: AG_NomeDia = "sexta"
        Case 7: AG_NomeDia = "sabado"
        Case Else: AG_NomeDia = CStr(d)
    End Select
End Function

Private Function AG_Ultima() As Variant
    On Error Resume Next
    AG_Ultima = ThisWorkbook.Worksheets(AG_ABA_PARAM).Range(AG_CEL_ULT).Value
    On Error GoTo 0
End Function

Private Sub AG_GravarUltima(ByVal quando As Date)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AG_ABA_PARAM)
    If ws Is Nothing Then Exit Sub
    If Trim$(CStr(ws.Range(AG_CEL_ROT).Value)) = "" Then
        ws.Range(AG_CEL_ROT).Value = "Cobranca semanal - ultimo envio automatico"
    End If
    ws.Range(AG_CEL_ULT).Value = quando
    ws.Range(AG_CEL_ULT).NumberFormat = "dd/mm/yyyy hh:mm"
    On Error GoTo 0
End Sub