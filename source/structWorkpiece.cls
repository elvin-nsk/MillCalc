VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructWorkpiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================

Public Shape As Shape
Public ElementsBunches As Collection

'===============================================================================

Public Property Get Self() As StructWorkpiece
  Set Self = Me
End Property
