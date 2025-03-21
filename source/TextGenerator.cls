VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Const LengthUnit As String = "м"
Const TimeUnit As String = "мин."
Const VelocityUnit As String = "м/мин."

Private Type typeThis
  Calculator As Calculator
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Calculator As Calculator) As TextGenerator
  With New TextGenerator
    Set Create = .Init(Calculator)
  End With
End Function

Friend Function Init(ByVal Calculator As Calculator) As TextGenerator
  With This
    Set .Calculator = Calculator
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get ToString(Optional ByVal ValuesSeparator As String = vbTab, _
                             Optional ByVal LinesSeparator As String = vbNewLine, _
                             Optional ByVal ParagraphsSeparator As String = vbNewLine & vbNewLine, _
                             Optional ByVal NumDigitsAfterDecimalToRound As Long = 0 _
                             ) As String
  With This.Calculator
    ToString = ToString & _
      "Время перемещения" & LinesSeparator & _
      "Вид фрезеровки" & LinesSeparator & _
      NamesAndMovementTimes(ValuesSeparator, LinesSeparator, NumDigitsAfterDecimalToRound) & _
      "ИТОГО:" & ValuesSeparator & VBA.Round(.TotalMovementTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      ParagraphsSeparator & _
      "Время фрезеровки" & LinesSeparator & _
      "Вид фрезеровки" & LinesSeparator & _
      NamesAndMillingTimes(ValuesSeparator, LinesSeparator, NumDigitsAfterDecimalToRound) & _
      "ИТОГО:" & ValuesSeparator & VBA.Round(.TotalMillingTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      ParagraphsSeparator & _
      "Снятие элементов" & ValuesSeparator & VBA.Round(.TotalRemovalTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      "Смена инструмента" & ValuesSeparator & VBA.Round(.TotalBeforeAndAfterProcessTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      "Подъём/опускание фрезы" & ValuesSeparator & VBA.Round(.TotalBeforeAndAfterElementTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      "Время установки заготовок" & ValuesSeparator & VBA.Round(.TotalInstallationTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator & _
      ParagraphsSeparator & _
      "ИТОГОВОЕ ВРЕМЯ:" & ValuesSeparator & ToHoursAndMinutes(.TotalTime)
  End With

End Property

Public Property Get Self() As TextGenerator
  Set Self = Me
End Property

'===============================================================================

Private Function ToHoursAndMinutes(ByVal Minutes As Long) As String
  ToHoursAndMinutes = VBA.Fix(Minutes / 60) & " час. " & Minutes Mod 60 & " мин."
End Function

Private Function NamesAndMovementTimes(ByVal ValuesSeparator As String, _
                                       ByVal LinesSeparator As String, _
                                       ByVal NumDigitsAfterDecimalToRound _
                                       ) As String
  Dim Process As structProcess
  For Each Process In This.Calculator.Processes
    With Process
      NamesAndMovementTimes = NamesAndMovementTimes & _
        .Name & ValuesSeparator & VBA.Round(.SummaryMovementTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator
    End With
  Next Process
End Function

Private Function NamesAndMillingTimes(ByVal ValuesSeparator As String, _
                                      ByVal LinesSeparator As String, _
                                      ByVal NumDigitsAfterDecimalToRound _
                                      ) As String
  Dim Process As structProcess
  For Each Process In This.Calculator.Processes
    With Process
      NamesAndMillingTimes = NamesAndMillingTimes & _
        .Name & ValuesSeparator & VBA.Round(.SummaryMillingTime, NumDigitsAfterDecimalToRound) & " " & TimeUnit & LinesSeparator
    End With
  Next Process
End Function
