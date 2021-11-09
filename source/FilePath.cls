VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IFilePath

'===============================================================================

Private Type typeThis
  Path As String
  NameWithoutExt As String
  Ext As String
End Type
Private This As typeThis

'===============================================================================

Public Function Create(Optional ByVal File As String) As IFilePath
  With New FilePath
    Set Create = .Init(File)
  End With
End Function

Friend Function Init(ByVal File As String) As IFilePath
  With This
    .Path = GetFilePath(File)
    .NameWithoutExt = GetFileNameWithoutExt(GetFileName(File))
    .Ext = GetFileExt(GetFileName(File))
  End With
  Set Init = Me
End Function

'===============================================================================

Private Property Get IFilePath_Ext() As String
  IFilePath_Ext = This.Ext
End Property
Private Property Let IFilePath_Ext(ByVal Value As String)
  This.Ext = Value
End Property

Private Property Get IFilePath_File() As String
  IFilePath_File = This.Path & This.NameWithoutExt & AddDotIfNotEmpty(This.Ext)
End Property

Private Property Get IFilePath_Name() As String
  IFilePath_Name = This.NameWithoutExt & AddDotIfNotEmpty(This.Ext)
End Property
Private Property Let IFilePath_Name(ByVal Value As String)
  With This
    .NameWithoutExt = GetFileNameWithoutExt(Value)
    .Ext = GetFileExt(Value)
  End With
End Property

Private Property Get IFilePath_NameWithoutExt() As String
  IFilePath_NameWithoutExt = This.NameWithoutExt
End Property
Private Property Let IFilePath_NameWithoutExt(ByVal Value As String)
  This.NameWithoutExt = Value
End Property

Private Property Get IFilePath_Path() As String
  IFilePath_Path = This.Path
End Property
Private Property Let IFilePath_Path(ByVal Value As String)
  This.Path = AddProperEndingToPath(Value)
End Property

Private Property Get IFilePath_Self() As IFilePath
  Set IFilePath_Self = Me
End Property

'===============================================================================

Private Function AddProperEndingToPath(ByVal Path As String) As String
  If Right$(Path, 1) <> "\" Then AddProperEndingToPath = Path & "\" _
  Else AddProperEndingToPath = Path
End Function

Private Function GetFileExt(FileName As String)
  GetFileExt = Right$(FileName, _
    Switch _
      (InStr(FileName, ".") = 0, _
        0, _
      InStr(FileName, ".") > 0, _
        Len(FileName) - InStrRev(FileName, ".")))
End Function

Private Function AddDotIfNotEmpty(ByVal Ext As String) As String
  If Ext = "" Then AddDotIfNotEmpty = "" Else AddDotIfNotEmpty = "." & Ext
End Function

Private Function GetFileName(ByVal FullFileName As String)
  GetFileName = Right$(FullFileName, Len(FullFileName) - InStrRev(FullFileName, "\"))
End Function

Private Function GetFileNameWithoutExt(ByVal FileName$) As String
  GetFileNameWithoutExt = Left$(FileName, _
    Switch _
      (InStr(FileName, ".") = 0, _
        Len(FileName), _
      InStr(FileName, ".") > 0, _
        InStrRev(FileName, ".") - 1))
End Function

Private Function GetFilePath(ByVal FullFileName As String)
  GetFilePath = Left$(FullFileName, InStrRev(FullFileName, "\"))
End Function
