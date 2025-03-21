VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  SizeRangeFrom As Double
  SizeRangeTo As Double
  MovementTime As Double
  RemovalTime As Double
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal SizeRangeFrom As Double, _
                       ByVal SizeRangeTo As Double, _
                       ByVal MovementTime As Double, _
                       ByVal RemovalTime As Double _
                       ) As ElementSize
  With New ElementSize
    Set Create = .Init(SizeRangeFrom, SizeRangeTo, MovementTime, RemovalTime)
  End With
End Function

Friend Function Init(ByVal SizeRangeFrom As Double, _
                     ByVal SizeRangeTo As Double, _
                     ByVal MovementTime As Double, _
                     ByVal RemovalTime As Double _
                     ) As ElementSize
  With This
    .SizeRangeFrom = SizeRangeFrom
    .SizeRangeTo = SizeRangeTo
    .MovementTime = MovementTime
    .RemovalTime = RemovalTime
  End With
  Set Init = Me
End Function

'===============================================================================

Public Property Get MovementTime() As Double
  MovementTime = This.MovementTime
End Property

Public Property Get RemovalTime() As Double
  RemovalTime = This.RemovalTime
End Property

Public Function MatchSize(ByVal Size As Double) As Boolean
  With This
    MatchSize = Size > .SizeRangeFrom And Size <= .SizeRangeTo
  End With
End Function

Public Property Get Self() As ElementSize
  Set Self = Me
End Property

'===============================================================================
