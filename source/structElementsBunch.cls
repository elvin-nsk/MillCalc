VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructElementsBunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================

Public Process As StructProcess
Public Elements As Collection

'===============================================================================

Public Property Get Self() As StructElementsBunch
  Set Self = Me
End Property
