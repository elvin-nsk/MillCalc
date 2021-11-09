VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "Model"
'@Interface
Option Explicit

'===============================================================================

Public Property Get Ext() As String
End Property
Public Property Let Ext(ByVal Value As String)
End Property

Public Property Get File() As String
End Property

Public Property Get Name() As String
End Property
Public Property Let Name(ByVal Value As String)
End Property

Public Property Get NameWithoutExt() As String
End Property
Public Property Let NameWithoutExt(ByVal Value As String)
End Property

Public Property Get Path() As String
End Property
Public Property Let Path(ByVal Value As String)
End Property

'===============================================================================

Public Property Get Self() As IFilePath
End Property
