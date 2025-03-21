VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Property Get Process() As structProcess
End Property

Public Property Get BBox() As Rect
End Property

Public Property Get MaxSize() As Double
End Property

Public Property Get Length() As Double
End Property

'===============================================================================

Public Property Get Self() As IElement
End Property
