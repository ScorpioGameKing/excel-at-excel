Attribute VB_Name = "LBActions"
Private ppt_app As PowerPoint.Application
Private ppt_presentation As PowerPoint.Presentation
Private template As String
Private file_name As String
Private file_path As String
Private init_run As Boolean

Public Sub INIT(temp As String, file As String, ext As String)
    If Not init_run Then
        MsgBox "Please wait a few seconds for the file to intialize. Thank You!", vbOKOnly, "LBActions INIT"
        Set ppt_app = CreateObject("Powerpoint.Application")
        template = temp
        SetFileName file, ext
        SetFilePath
        init_run = True
        MsgBox "Intialization Complete. Please Reload the file if issues occur. Thank You!", vbOKOnly, "LBActions INIT"
    End If
End Sub

Sub SetFileName(name As String, ext As String)
    file_name = name & " " & Replace(Date, "/", "-") & ext
End Sub

Sub SetFilePath()
    file_path = Application.ActiveWorkbook.path & "/"
End Sub

Sub OpenPPT()
    If ppt_app.Presentations.Count > 0 Then
        Set ppt_presentation = FindOpenPPT
    Else
        Set ppt_presentation = FindSavedPPT
    End If
End Sub

Sub UpdatePPT()
    OpenPPT
    ClearSlideOf 1, "PPHBoard"
    ClearSlideOf 2, "PPHBoard"
    ClearSlideOf 3, "PPHBoard"
    PasteOnSlide 1, "Leaderboard 1", "A1", "B6", "P"
    PasteOnSlide 2, "Leaderboard 2", "A1", "B6", "P"
    PasteOnSlide 3, "Leaderboard 3", "A1", "B6", "P"
End Sub

Function FindOpenPPT() As PowerPoint.Presentation
    Dim open_file As Boolean: open_file = False
    For Each p In ppt_app.Presentations
        If p.FullName = file_path & file_name Then
            Set FindOpenPPT = p
            open_file = True
        End If
    Next
    If Not open_file Then Set FindOpenPPT = FindSavedPPT
End Function

Function FindSavedPPT() As PowerPoint.Presentation
    On Error GoTo NoFile
    Set FindSavedPPT = ppt_app.Presentations.Open(file_path & file_name)
    Exit Function
NoFile:
    Set FindSavedPPT = CreateTemplatedPPT
End Function

Function CreateTemplatedPPT() As PowerPoint.Presentation
    Dim temp_ppt As PowerPoint.Presentation: Set temp_ppt = ppt_app.Presentations.Open(file_path & template)
    temp_ppt.SaveAs (file_path & file_name)
    Set CreateTemplatedPPT = temp_ppt
End Function

Function FindCaptureArea(sheet As String, topLeft As String, innerLeft As String, bottomRight As String) As String
    FindCaptureArea = topLeft & ":" & bottomRight & Sheets(sheet).Range(innerLeft).End(xlDown).Row
End Function

Sub ClearSlideOf(index As Integer, name As String)
    For Each s In ppt_presentation.Slides(index).Shapes
        If s.name = name Then
            s.Delete
        End If
    Next
End Sub

Sub PasteOnSlide(index As Integer, sheet As String, topLeft As String, innerLeft As String, bottomRight As String)
    Dim cap_area As Range: Set cap_area = Sheets(sheet).Range(FindCaptureArea(sheet, topLeft, innerLeft, bottomRight))
    cap_area.CopyPicture xlScreen, xlBitmap
    With ppt_presentation.Slides(index).Shapes.PasteSpecial(ppPasteBitmap)
        .name = "PPHBoard"
    End With
End Sub
