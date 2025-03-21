VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapesParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Shapes As ShapeRange
  WorkpieceColors As Collection
  Processes As Collection
  Workpieces As Collection
  Elements As Collection
  InvalidShapes As ShapeRange
  ElementsOutsideWorkpieces As Collection
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Shapes As ShapeRange, _
                       ByVal WorkpieceColors As Collection, _
                       ByVal Processes As Collection) As ShapesParser
  With New ShapesParser
    Set Create = .Init(Shapes, WorkpieceColors, Processes)
  End With
End Function

Friend Function Init(ByVal Shapes As ShapeRange, _
                     ByVal WorkpieceColors As Collection, _
                     ByVal Processes As Collection) As ShapesParser
  With This
    Set .Shapes = Shapes
    Set .WorkpieceColors = WorkpieceColors
    Set .Processes = Processes
    Set .Workpieces = New Collection
    Set .Elements = New Collection
    Set .InvalidShapes = CreateShapeRange
    Set .ElementsOutsideWorkpieces = New Collection
  End With
  ParseShapes
  SeparateByWorkpieces
  Set Init = Me
End Function

'===============================================================================

Public Property Get Workpieces() As Collection
  Set Workpieces = This.Workpieces
End Property

Public Property Get InvalidShapes() As ShapeRange
  Set InvalidShapes = This.InvalidShapes
End Property

Public Property Get ElementsOutsideWorkpieces() As Collection
  Set ElementsOutsideWorkpieces = This.ElementsOutsideWorkpieces
End Property

Public Property Get Self() As ShapesParser
  Set Self = Me
End Property

'===============================================================================

Private Sub ParseShapes()
  With This
    Dim Shape As Shape
    For Each Shape In .Shapes.Shapes.FindShapes
      ProcessShape Shape
    Next Shape
  End With
End Sub

Private Sub SeparateByWorkpieces()
  With This
    Dim Element As IElement
    For Each Element In .Elements
      SortElementIntoWorkpiece Element
    Next Element
  End With
End Sub

'-------------------------------------------------------------------------------

Private Sub ProcessShape(ByVal Shape As Shape)
  If ProcessShapeAsWorkpiece(Shape) Then Exit Sub
  If ProcessShapeAsElementSubPath(Shape) Then Exit Sub
End Sub

Private Function ProcessShapeAsWorkpiece(ByVal Shape As Shape) As Boolean
  With This
    If Not ShapeIsVaidWorkpiece(Shape) Then Exit Function
    If ColorMatchCollection(Shape.Outline.Color, .WorkpieceColors) Then
      With New structWorkpiece
        Set .Shape = Shape
        Set .ElementsBunches = New Collection
        This.Workpieces.Add .Self
      End With
      ProcessShapeAsWorkpiece = True
      Exit Function
    End If
  End With
End Function

Private Function ProcessShapeAsElementSubPath(ByVal Shape As Shape) As Boolean
  With This
    If Not ShapeIsVaidElement(Shape) Then
      .InvalidShapes.Add Shape
      Exit Function
    End If
    Dim Process As structProcess
    For Each Process In .Processes
      If ColorMatchCollection(Shape.Outline.Color, Process.Colors) Then
        AddToElementsAsElementSubPath Shape, Process
        ProcessShapeAsElementSubPath = True
        Exit Function
      End If
    Next Process
  End With
End Function

Private Sub AddToElementsAsElementSubPath(ByVal Shape As Shape, _
                                          ByVal Process As structProcess)
  With This
    Dim SubPath As SubPath
    For Each SubPath In Shape.Curve.SubPaths
      .Elements.Add ElementSubPath.Create(SubPath, Process)
    Next SubPath
  End With
End Sub

Private Function ColorMatchCollection(ByVal Color As Color, _
                                      ByVal ColorsCollection As Collection _
                                      ) As Boolean
  Dim tempColor As Color
  For Each tempColor In ColorsCollection
    If tempColor.IsSame(Color) Then
      ColorMatchCollection = True
      Exit Function
    End If
  Next tempColor
End Function

Private Function ShapeIsVaidWorkpiece(ByVal Shape As Shape) As Boolean
  On Error GoTo Catch
  If Shape.Outline.Type = cdrNoOutline Then Exit Function
  ShapeIsVaidWorkpiece = True
Catch:
  On Error GoTo 0
End Function

Private Function ShapeIsVaidElement(ByVal Shape As Shape) As Boolean
  On Error GoTo Catch
  If Shape.Outline.Type = cdrNoOutline Then Exit Function
  If Shape.Curve Is Nothing Then Exit Function
  If Shape.Curve.SubPaths.Count = 0 Then Exit Function
  If Shape.Curve.Length = 0 Then Exit Function
  ShapeIsVaidElement = True
Catch:
  On Error GoTo 0
End Function

Private Sub SortElementIntoWorkpiece(ByVal Element As IElement)
  With This
    Dim Workpiece As structWorkpiece
    For Each Workpiece In .Workpieces
      If IsOverlaps(Element.BBox, Workpiece.Shape.BoundingBox) Then
        FindOrAddElementsBunch(Element, Workpiece).Elements.Add Element
        Exit Sub
      End If
    Next Workpiece
    .ElementsOutsideWorkpieces.Add Element
  End With
End Sub

Private Function FindOrAddElementsBunch(ByVal Element As IElement, _
                                        ByVal Workpiece As structWorkpiece _
                                        ) As structElementsBunch
  Dim Bunch As structElementsBunch
  For Each Bunch In Workpiece.ElementsBunches
    If Bunch.Process Is Element.Process Then
      Set FindOrAddElementsBunch = Bunch
      Exit Function
    End If
  Next Bunch
  Set FindOrAddElementsBunch = CreateElementsBunch
  Set FindOrAddElementsBunch.Process = Element.Process
  Workpiece.ElementsBunches.Add FindOrAddElementsBunch
End Function

Private Function CreateElementsBunch() As structElementsBunch
  With New structElementsBunch
    Set .Elements = New Collection
    Set CreateElementsBunch = .Self
  End With
End Function

Private Function IsOverlaps(ByVal Rect1 As Rect, _
                            ByVal Rect2 As Rect _
                            ) As Boolean
  Dim BBox1 As Rect, BBox2 As Rect
  Set BBox1 = Rect1.GetCopy
  Set BBox2 = Rect2.GetCopy
  SlightlyInflateRectIfSideIsZero BBox1
  SlightlyInflateRectIfSideIsZero BBox2
  IsOverlaps = Not BBox1.Intersect(BBox2).IsEmpty
End Function

Private Sub SlightlyInflateRectIfSideIsZero(ByVal Rect As Rect)
  Const Ratio As Double = 1000
  If Rect.Width = 0 Then
    Rect.Width = Rect.Height / Ratio
  ElseIf Rect.Height = 0 Then
    Rect.Height = Rect.Width / Ratio
  End If
End Sub
