Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

Private Sub Workbook_Open()
    ' Ao abrir, busca o dia e a hora configurados e faz o agendamento
    Verificar_E_Agendar_Envio
End Sub