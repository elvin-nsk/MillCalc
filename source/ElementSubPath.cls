VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementSubPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IElement

'===============================================================================

Private Type This
  Process As StructProcess
  BBox As Rect
  MaxSize As Double
  Length As Double
End Type
Private This As This

'===============================================================================

Public Function Create(ByVal SubPath As SubPath, _
                       ByVal Process As StructProcess _
                       ) As IElement
  With New ElementSubPath
    Set Create = .Init(SubPath, Process)
  End With
End Function

Friend Function Init(ByVal SubPath As SubPath, _
                     ByVal Process As StructProcess _
                     ) As IElement
  With This
    Set .Process = Process
  End With
  Parse SubPath
  Set Init = Me
End Function

'===============================================================================

Private Property Get IElement_Process() As StructProcess
  Set IElement_Process = This.Process
End Property

Private Property Get IElement_BBox() As Rect
 Set IElement_BBox = This.BBox
End Property

Private Property Get IElement_MaxSize() As Double
  IElement_MaxSize = This.MaxSize
End Property

Private Property Get IElement_Length() As Double
  IElement_Length = This.Length
End Property

Private Property Get IElement_Self() As IElement
  Set IElement_Self = Me
End Property

'===============================================================================

Private Sub Parse(ByVal SubPath As SubPath)
  With This
    Set .BBox = SubPath.BoundingBox.GetCopy
    CalcMaxSize SubPath
    .Length = SubPath.Length
  End With
End Sub

Private Sub CalcMaxSize(ByVal SubPath As SubPath)
  If SubPath.SizeWidth > SubPath.SizeHeight Then
    This.MaxSize = SubPath.SizeWidth
  Else
    This.MaxSize = SubPath.SizeHeight
  End If
End Sub
