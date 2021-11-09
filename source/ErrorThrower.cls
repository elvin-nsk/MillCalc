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
  Throw "Не найден файл " & vbCrLf & _
  File
End Sub

Public Sub ErrorReadingFile(ByVal FileName As String)
  Throw "Ошибка чтения файла " & FileName
End Sub

Public Sub ErrorCSV(ByVal FileName As String)
  Throw "Ошибка структуры таблицы " & FileName
End Sub

Public Sub ErrorCSVSize(ByVal FileName As String)
  Throw "Мало строк в таблице " & FileName
End Sub

Public Sub ErrorRecord(ByVal FileName As String, _
                             ByVal Row As Long)
  Throw "Неверный тип данных в таблице" & vbCrLf & _
        "Файл: " & FileName & vbCrLf & _
        "Строка: " & VBA.CStr(Row)
End Sub

Public Sub ErrorElementSize(ByVal FileName As String, _
                            ByVal Row1 As Long)
  Throw "Размер элемента в таблице " & FileName & vbCrLf & _
        "на строке " & Row1 & " больше, чем на следующий строке"
End Sub

Public Sub ErrorColorInCSV(ByVal FileName As String, _
                             ByVal Row As Long, _
                             ByVal Col As Long)
  Throw "Неверно задан цвет" & vbCrLf & _
        "Файл: " & FileName & vbCrLf & _
        "Строка: " & VBA.CStr(Row) & vbCrLf & _
        "Столбец: " & VBA.CStr(Col)
End Sub

Public Sub ErrorColorInINI(ByVal FileName As String)
  Throw "Неверно задан цвет в файле " & FileName
End Sub

Public Sub Throw(ByVal Message As String)
  VBA.Err.Raise CustomError, "ErrorHandler", Message
End Sub
