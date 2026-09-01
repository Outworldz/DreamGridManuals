Attribute VB_Name = "Module1"

Sub BatchExportManualsToFilteredHtml()

    Dim sourceFolder As String
    Dim destFolder As String
    Dim fileName As String
    Dim baseName As String
    Dim doc As Document
    Dim wasAlreadyOpen As Boolean
    Dim count As Long
    Dim errorList As String

    sourceFolder = "O:\Opensim\!Manuals\"
    destFolder = "O:\Opensim\!TandyV7\OutworldzFiles\Help\"

    If Dir(destFolder, vbDirectory) = "" Then
        MkDir destFolder
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    fileName = Dir(sourceFolder & "*.docx")

    Do While fileName <> ""

        If LCase(fileName) = LCase("DreamGrid Manual.docx") Then
            fileName = Dir()
            GoTo ContinueLoop
        End If

        baseName = Left(fileName, InStrRev(fileName, ".") - 1)

        wasAlreadyOpen = False
        Dim d As Document
        For Each d In Documents
            If d.FullName = sourceFolder & fileName Then
                Set doc = d
                wasAlreadyOpen = True
                Exit For
            End If
        Next d

        If Not wasAlreadyOpen Then
            Set doc = Documents.Open(fileName:=sourceFolder & fileName, ReadOnly:=True, AddToRecentFiles:=False)
        End If

        On Error Resume Next
        Err.Clear
        doc.SaveAs2 fileName:=destFolder & baseName & ".htm", _
            FileFormat:=wdFormatFilteredHTML

        If Err.Number <> 0 And Err.Number <> 5153 Then
            errorList = errorList & fileName & ": Error " & Err.Number & " - " & Err.Description & vbCrLf
        End If
        On Error GoTo 0

        If Not wasAlreadyOpen Then
            doc.Close SaveChanges:=wdDoNotSaveChanges
        End If

        count = count + 1
        fileName = Dir()

ContinueLoop:
    Loop

    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    If errorList = "" Then
        MsgBox count & " file(s) processed to Web Page, Filtered in:" & vbCrLf & destFolder
    Else
        MsgBox count & " file(s) processed. Errors encountered:" & vbCrLf & errorList
    End If

End Sub

