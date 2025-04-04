VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' Config
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const iniFileName As String = "elvin_MillCalc.ini"

Private Type This
  SourceDataFolder As String
  ConstantsFileName As String
  ProcessesFileName As String
  ElementSizesFileName As String
End Type
Private This As This

'===============================================================================

Public WorkpieceColors As Collection
Public WorkpieceInstallationTime As Double
Public WorkpieceRemovalTimeMod As Double
Public BeforeProcessTime As Double
Public AfterProcessTime As Double
Public BeforeElementTime As Double
Public AfterElementTime As Double
Public ElementRemovalTimeMod As Double

Public Processes As Collection
Public ElementSizes As Collection

'===============================================================================

Public Function CreateAndLoad() As Config
  Set CreateAndLoad = New Config
  CreateAndLoad.Load
End Function

'===============================================================================

Public Sub Load()
  Assign iniRead
  ReadConstants
  ReadProcesses
  ReadElementSizes
End Sub

Public Sub Save()
  Assign iniWrite
End Sub

Public Property Get Self() As Config
  Set Self = Me
End Property

'===============================================================================

Private Sub Assign(ByVal Action As iniAction)

  With iniFile.Create
  
    iniFilePrepare .Self, Action
    
    .KeyVar "SourceDataFolder", This.SourceDataFolder, "c:\temp\SourceData\"
    .KeyVar "ConstantsFileName", This.ConstantsFileName, "Constants.ini"
    .KeyVar "ProcessesFileName", This.ProcessesFileName, "Processes.csv"
    .KeyVar "ElementSizesFileName", This.ElementSizesFileName, "ElementSizes.csv"
  
  End With
  
End Sub

Private Sub ReadConstants()

  Dim File As FileSpec
  Set File = GetFilePathFromDataFolder(This.ConstantsFileName)
  If Not FileExist(File) Then ErrorThrower.ErrorFileNotFound File

  With iniFile.Create
    
    .File = File
    .Action = iniRead
    
    .Section = "Workpiece"
    AddColorToWorkpieceColors .Self
    .KeyVar "InstallationTime", WorkpieceInstallationTime
    .KeyVar "RemovalTimeMod", WorkpieceRemovalTimeMod
    
    .Section = "Process"
    .KeyVar "BeforeTime", BeforeProcessTime
    .KeyVar "AfterTime", AfterProcessTime
    
    .Section = "Element"
    .KeyVar "BeforeTime", BeforeElementTime
    .KeyVar "AfterTime", AfterElementTime
    .KeyVar "RemovalTimeMod", ElementRemovalTimeMod
    
  End With

End Sub

Private Sub ReadProcesses()

  Dim File As FileSpec
  Set File = GetFilePathFromDataFolder(This.ProcessesFileName)
  
  Dim Csv As Collection
  Set Csv = ReadCSV(File)
  
  If Csv.Count < 2 Then ErrorThrower.ErrorCSVSize File.Name
      
  Set Processes = New Collection
  On Error GoTo Catch
  Dim i
  For i = 2 To Csv.Count
    Processes.Add ReadProcessesRecord(Csv(i))
  Next i
  On Error GoTo 0
  
  Exit Sub

Catch:
  ErrorThrower.ErrorRecord File.Name, i

End Sub

Private Sub ReadElementSizes()

  Dim File As FileSpec
  Set File = GetFilePathFromDataFolder(This.ElementSizesFileName)
  
  Dim Csv As Collection
  Set Csv = ReadCSV(File)
  
  If Csv.Count < 2 Then ErrorThrower.ErrorCSVSize File.Name
  
  Set ElementSizes = New Collection
  On Error GoTo Catch
  Dim i
  For i = 2 To Csv.Count
    ElementSizes.Add ReadElementSizesRecord(i, Csv)
  Next i
  On Error GoTo 0
  
  Exit Sub

Catch:
  ErrorThrower.ErrorRecord File.Name, i
  
End Sub

Private Sub AddColorToWorkpieceColors(ByVal File As iniFile)
  Set WorkpieceColors = New Collection
  Dim ColorString As String
  On Error GoTo Catch
  With File
    Dim i As Long
    For i = 1 To 999
      If .Entry(iniRead, "Workpiece", "Color" & VBA.CStr(i), ColorString) Then
        WorkpieceColors.Add CreateColor(ColorString)
      Else
        Exit For
      End If
    Next i
  End With
  On Error GoTo 0
  Exit Sub
Catch:
  ErrorThrower.ErrorColorInINI This.ConstantsFileName
End Sub

Private Function ReadProcessesRecord(ByVal Record As Collection) As StructProcess
  With New StructProcess
    .Name = Record(1)
    .Removal = VBA.CBool(Record(2))
    .Velocity = VBA.CDbl(Record(3))
    Set .Colors = New Collection
    Dim i
    For i = 4 To Record.Count
      .Colors.Add CreateColor(Record(i))
    Next i
    Set ReadProcessesRecord = .Self
  End With
End Function

Private Function ReadElementSizesRecord(ByVal Index As Long, _
                                        ByVal Csv As Collection _
                                        ) As ElementSize
  Dim NextSize As Double
  If Csv.Count < (Index + 1) Then NextSize = 2147483647 Else NextSize = Csv(Index + 1)(1)
  If Csv(Index)(1) > NextSize Then _
    ErrorThrower.ErrorElementSize This.ElementSizesFileName, Index
  Set ReadElementSizesRecord = _
      ElementSize.Create(Csv(Index)(1), NextSize, Csv(Index)(2), Csv(Index)(3))
End Function

Private Function ReadCSV(ByVal File As FileSpec) As Collection
  On Error GoTo Catch
  With CsvUtils.Create(";")
    Set ReadCSV = .ParseCsvToCollection(LoadStrFromFile(File), True)
  End With
  On Error GoTo 0
  Exit Function
Catch:
  ErrorThrower.ErrorCSV File.Name
End Function

Private Function GetFilePathFromDataFolder(ByVal FileName As String) As FileSpec
  With FileSpec.New_
    .Name = FileName
    .Path = This.SourceDataFolder
    Set GetFilePathFromDataFolder = .Self
  End With
End Function

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

  Dim iniFilePath As String

  With iniFile
    Select Case True
    Case FileExist(GMSManager.UserGMSPath & iniFileName)
      .File = Application.GMSManager.UserGMSPath & iniFileName
      .Action = Action
    Case FileExist(GMSManager.GMSPath & iniFileName)
      .File = Application.GMSManager.GMSPath & iniFileName
      .Action = Action
    Case Else
      .File = GMSManager.UserGMSPath & iniFileName
      If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
    End Select
  End With

End Sub

Private Function FileExist(ByVal File As String) As Boolean
  If VBA.Len(VBA.Dir(File)) > 0 Then
    FileExist = True
  End If
End Function

Private Function LoadStrFromFile(ByVal File As String) As String
  Dim FileNum&: FileNum = VBA.FreeFile
  On Error GoTo Catch
  Open File For Input As #FileNum
  LoadStrFromFile = Input(LOF(FileNum), FileNum)
  Close #FileNum
  On Error GoTo 0
  Exit Function
Catch:
  ErrorThrower.ErrorReadingFile FileSpec.New_(File).Name
End Function
