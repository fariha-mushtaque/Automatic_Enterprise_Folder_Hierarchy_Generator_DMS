Attribute VB_Name = "Module2"
Option Explicit

Private ProcessedFolders As Object

Private Const MAX_LEVELS As Integer = 8

Private RootFolder As String

Private Levels(1 To MAX_LEVELS) As String

Private TerminalPath As String

Private CreatedCount As Long
Private ExistingCount As Long
Private ErrorCount As Long
Private StartTime As Double


Public Sub CreateFolderHierarchy()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = Worksheets("Hierarchy")

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Root Folder"

        If .Show <> -1 Then Exit Sub

        RootFolder = .SelectedItems(1)

    End With

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    CreatedCount = 0
    ExistingCount = 0
    ErrorCount = 0
    
    Set ProcessedFolders = CreateObject("Scripting.Dictionary")

    StartTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Building Folder Hierarchy..."

    For r = 2 To LastRow

        If Not RowIsBlank(ws, r) Then

    ReadCurrentRow ws, r

    BuildCurrentPath

    If IsTerminal(ws, r, LastRow) Then

        CreateTemplateFolders TerminalPath

    End If

End If
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True
Dim ElapsedTime As String

ElapsedTime = Format((Timer - StartTime) / 86400, "hh:mm:ss")


MsgBox _
"Folder Hierarchy Creation Completed Successfully!" & vbCrLf & vbCrLf & _
"Destination :" & vbCrLf & RootFolder & vbCrLf & vbCrLf & _
"New Folders Created : " & CreatedCount & vbCrLf & _
"Existing Folders : " & ExistingCount & vbCrLf & _
"Errors : " & ErrorCount & vbCrLf & _
"Execution Time : " & ElapsedTime, _
vbInformation, "Execution Summary"

End Sub

Private Sub ReadCurrentRow(ws As Worksheet, ByVal rw As Long)

    Dim i As Integer

    For i = 1 To MAX_LEVELS

        If Trim(ws.Cells(rw, i).Value) <> "" Then

            Levels(i) = CleanFolderName(ws.Cells(rw, i).Value)

            Dim j As Integer

            For j = i + 1 To MAX_LEVELS
                Levels(j) = ""
            Next j

        End If

    Next i

End Sub

Private Function CleanFolderName(txt As String) As String

    txt = Trim(txt)

    txt = Replace(txt, "/", "-")
    txt = Replace(txt, "\", "-")
    txt = Replace(txt, ":", "-")
    txt = Replace(txt, "*", "")
    txt = Replace(txt, "?", "")
    txt = Replace(txt, """", "")
    txt = Replace(txt, "<", "")
    txt = Replace(txt, ">", "")
    txt = Replace(txt, "|", "")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    CleanFolderName = txt

End Function

Private Sub BuildCurrentPath()

    Dim i As Integer
    Dim CurrentPath As String

    CurrentPath = RootFolder
    TerminalPath = ""

    For i = 1 To MAX_LEVELS

        If Levels(i) <> "" Then

            CurrentPath = CurrentPath & "\" & Levels(i)

            EnsureFolder CurrentPath

            TerminalPath = CurrentPath

        End If

    Next i

End Sub

Private Sub EnsureFolder(ByVal FolderPath As String)

    On Error GoTo ErrHandler

    'Already handled during this run
    If ProcessedFolders.Exists(UCase(FolderPath)) Then Exit Sub

    ProcessedFolders.Add UCase(FolderPath), True

    If Dir(FolderPath, vbDirectory) = "" Then

        MkDir FolderPath

        CreatedCount = CreatedCount + 1

    Else

        ExistingCount = ExistingCount + 1

    End If

    Exit Sub

ErrHandler:

    ErrorCount = ErrorCount + 1

End Sub
Private Function IsTerminal(ws As Worksheet, _
                            ByVal rw As Long, _
                            ByVal LastRow As Long) As Boolean

    Dim CurrentLevel As Integer
    Dim i As Integer

    'Last row is always terminal
    If rw = LastRow Then
        IsTerminal = True
        Exit Function
    End If

    'Find current deepest level
    CurrentLevel = 0

    For i = MAX_LEVELS To 1 Step -1

        If Trim(ws.Cells(rw, i).Value) <> "" Then
            CurrentLevel = i
            Exit For
        End If

    Next i

    'If next row contains data one level deeper,
    'current row is NOT terminal
    If CurrentLevel < MAX_LEVELS Then

        If Trim(ws.Cells(rw + 1, CurrentLevel + 1).Value) <> "" Then

            IsTerminal = False
            Exit Function

        End If

    End If

    IsTerminal = True

End Function

Private Sub CreateTemplateFolders(ByVal TargetFolder As String)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim ParentFolder As String
    Dim ChildFolder As String

    Dim ParentPath As String
    Dim ChildPath As String

    Set ws = Worksheets("Templates")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To LastRow

        ParentFolder = CleanFolderName(Trim(ws.Cells(r, 1).Value))
        ChildFolder = CleanFolderName(Trim(ws.Cells(r, 2).Value))

        If ParentFolder <> "" Then

            ParentPath = TargetFolder & "\" & ParentFolder

            EnsureFolder ParentPath

            If ChildFolder <> "" Then

                ChildPath = ParentPath & "\" & ChildFolder

                EnsureFolder ChildPath

            End If

        End If

    Next r

End Sub

Private Function RowIsBlank(ws As Worksheet, ByVal rw As Long) As Boolean

    Dim i As Integer

    RowIsBlank = True

    For i = 1 To MAX_LEVELS

        If Trim(ws.Cells(rw, i).Value) <> "" Then

            RowIsBlank = False
            Exit Function

        End If

    Next i

End Function


