Attribute VB_Name = "MillCalc"
'===============================================================================
' Макрос           : MillCalc
' Версия           : 2025.04.04
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "MillCalc"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2025.04.04"

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo Catch
  
  Dim Shapes As ShapeRange
  If Not InputData.ExpectShapes.Ok(Shapes) Then Exit Sub
  
  ActiveDocument.Unit = cdrMeter
  
  Dim Cfg As Config
  Set Cfg = Config.CreateAndLoad
  Dim Parser As ShapesParser
  Set Parser = ShapesParser.Create(Shapes, Cfg.WorkpieceColors, Cfg.Processes)
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
                      ByVal Parser As ShapesParser)
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
