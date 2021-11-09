VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorThrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Singleton
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Sub Class_Initialize()
  If Not Me Is ErrorThrower Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Sub ErrorFileNotFound(ByVal File As String)
  Throw "�� ������ ���� " & vbCrLf & _
  File
End Sub

Public Sub ErrorReadingFile(ByVal FileName As String)
  Throw "������ ������ ����� " & FileName
End Sub

Public Sub ErrorCSV(ByVal FileName As String)
  Throw "������ ��������� ������� " & FileName
End Sub

Public Sub ErrorCSVSize(ByVal FileName As String)
  Throw "���� ����� � ������� " & FileName
End Sub

Public Sub ErrorRecord(ByVal FileName As String, _
                             ByVal Row As Long)
  Throw "�������� ��� ������ � �������" & vbCrLf & _
        "����: " & FileName & vbCrLf & _
        "������: " & VBA.CStr(Row)
End Sub

Public Sub ErrorElementSize(ByVal FileName As String, _
                            ByVal Row1 As Long)
  Throw "������ �������� � ������� " & FileName & vbCrLf & _
        "�� ������ " & Row1 & " ������, ��� �� ��������� ������"
End Sub

Public Sub ErrorColorInCSV(ByVal FileName As String, _
                             ByVal Row As Long, _
                             ByVal Col As Long)
  Throw "������� ����� ����" & vbCrLf & _
        "����: " & FileName & vbCrLf & _
        "������: " & VBA.CStr(Row) & vbCrLf & _
        "�������: " & VBA.CStr(Col)
End Sub

Public Sub ErrorColorInINI(ByVal FileName As String)
  Throw "������� ����� ���� � ����� " & FileName
End Sub

Public Sub Throw(ByVal Message As String)
  VBA.Err.Raise CustomError, "ErrorHandler", Message
End Sub