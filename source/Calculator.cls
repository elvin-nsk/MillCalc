VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type This
  ShapesParser As ShapesParser
  Config As Config
  Processes As Collection
  TotalMillingTime As Double
  TotalRemovalTime As Double
  TotalBeforeAndAfterProcessTime As Double
  TotalBeforeAndAfterElementTime As Double
  TotalInstallationTime As Double
  TotalTime As Double
End Type
Private This As This

'===============================================================================

Public Function Create(ByVal ShapesParser As ShapesParser, _
                       ByVal Config As Config) As Calculator
  With New Calculator
    Set Create = .Init(ShapesParser, Config)
  End With
End Function

Friend Function Init(ByVal ShapesParser As ShapesParser, _
                     ByVal Config As Config) As Calculator
  With This
    Set .ShapesParser = ShapesParser
    Set .Config = Config
    Set .Processes = Config.Processes
  End With
  Calc
  Set Init = Me
End Function

'===============================================================================

Public Property Get Processes() As Collection
  Set Processes = This.Processes
End Property

Public Property Get TotalMillingTime() As Double
  TotalMillingTime = This.TotalMillingTime
End Property

Public Property Get TotalRemovalTime() As Double
  TotalRemovalTime = This.TotalRemovalTime
End Property

Public Property Get TotalBeforeAndAfterProcessTime() As Double
  TotalBeforeAndAfterProcessTime = This.TotalBeforeAndAfterProcessTime
End Property

Public Property Get TotalBeforeAndAfterElementTime() As Double
  TotalBeforeAndAfterElementTime = This.TotalBeforeAndAfterElementTime
End Property

Public Property Get TotalInstallationTime() As Double
  TotalInstallationTime = This.TotalInstallationTime
End Property

Public Property Get TotalTime() As Double
  TotalTime = This.TotalTime
End Property

Public Property Get Self() As Calculator
  Set Self = Me
End Property

'===============================================================================

Private Sub Calc()
  
  Dim Workpiece As StructWorkpiece
  Dim Bunch As StructElementsBunch
  Dim Element As IElement
  
  With This
  
    For Each Workpiece In .ShapesParser.Workpieces
      PerWorkpiece
      For Each Bunch In Workpiece.ElementsBunches
        PerProcess
        For Each Element In Bunch.Elements
          PerElement Element
        Next Element
      Next Bunch
    Next Workpiece
  
  .TotalTime = _
    .TotalMillingTime _
  + .TotalRemovalTime _
  + .TotalBeforeAndAfterProcessTime _
  + .TotalBeforeAndAfterElementTime _
  + .TotalInstallationTime
  
  End With

End Sub

Private Sub PerElement(ByVal Element As IElement)
  Dim Size As ElementSize
  Set Size = SizeOf(Element)
  Dim MillingTime As Double
  With This
    MillingTime = Element.Length / Element.Process.Velocity
    Element.Process.SummaryMillingTime = Element.Process.SummaryMillingTime _
      + MillingTime
    .TotalMillingTime = .TotalMillingTime + MillingTime
    If Element.Process.Removal Then _
      .TotalRemovalTime = .TotalRemovalTime _
        + Size.RemovalTime _
        + .Config.ElementRemovalTimeMod
    .TotalBeforeAndAfterElementTime = .TotalBeforeAndAfterElementTime _
      + .Config.BeforeElementTime _
      + .Config.AfterElementTime
  End With
End Sub

Private Sub PerWorkpiece()
  With This
    .TotalInstallationTime = .TotalInstallationTime _
      + .Config.WorkpieceInstallationTime
    .TotalRemovalTime = .TotalRemovalTime _
      + .Config.WorkpieceRemovalTimeMod
  End With
End Sub

Private Sub PerProcess()
  With This
    .TotalBeforeAndAfterProcessTime = .TotalBeforeAndAfterProcessTime _
      + .Config.BeforeProcessTime _
      + .Config.AfterProcessTime
  End With
End Sub

Private Function SizeOf(ByVal Element As IElement) As ElementSize
  Dim Size As ElementSize
  For Each Size In This.Config.ElementSizes
    If Size.MatchSize(Element.MaxSize) Then
      Set SizeOf = Size
      Exit Function
    End If
  Next Size
End Function
