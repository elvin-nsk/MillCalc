VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================

Public Name As String
Public Removal As Boolean
Public Velocity As Double
Public Colors As Collection
Public SummaryMillingTime As Double
Public SummaryMovementTime As Double

'===============================================================================

Public Property Get Self() As StructProcess
  Set Self = Me
End Property
