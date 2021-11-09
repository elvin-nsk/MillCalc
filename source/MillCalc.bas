Attribute VB_Name = "MillCalc"
'===============================================================================
' Макрос           : MillCalc
' Версия           : 2021.07.14
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then
    MsgBox "Нет активного документа"
    Exit Sub
  End If
  
  ActiveDocument.Unit = cdrMeter
  
  Dim Cfg As Config
  Set Cfg = Config.CreateAndLoad
  Dim Parser As PageParser
  Set Parser = PageParser.Create(ActivePage, Cfg.WorkpieceColors, Cfg.Processes)
  Dim Calc As Calculator
  Set Calc = Calculator.Create(Parser, Cfg)
  Dim Text As TextGenerator
  Set Text = TextGenerator.Create(Calc)
  
  Presenter Text, Parser
  
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"

End Sub

Private Sub Presenter(ByVal Text As TextGenerator, _
                      ByVal Parser As PageParser)
  With New MainView
    If Parser.ElementsOutsideWorkpieces.Count > 0 Then
      .lbOutsideElements = _
      Parser.ElementsOutsideWorkpieces.Count & " элементов за пределами заготовок"
      .lbOutsideElements.Visible = True
    End If
    .Output.Value = Text.ToString
    .Show
  End With
End Sub
