Attribute VB_Name = "RealityCheck"
'----- Suite of Exists, Contains, Find, Match, etc style Functions

'----- Module Checker
'----- REQUIRES: Microsoft Visual Basic For Applications Extensibility 5.3
Public Function IsModuleLoaded(name As String, Optional wb As Workbook) As Boolean
    Dim j As Long
    Dim vbcomp As VBComponent
    Dim modules As New Collection
    IsModuleLoaded = False

    '----- Check if values are set
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If (name = "") Then
        GoTo errorname
    End If

    '----- Collect Class Module Names
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If ((vbcomp.Type = vbext_ct_StdModule) Or (vbcomp.Type = vbext_ct_ClassModule)) Then
            modules.Add vbcomp.name
        End If
    Next vbcomp

    '----- Compare for the target
    For j = 1 To modules.Count
        If (name = modules.Item(j)) Then
            IsModuleLoaded = True
        End If
    Next j
    j = 0

    '----- Missing Jump
    If (IsModuleLoaded = False) Then
        GoTo notfound
    End If

    '----- Error Handling
    If (0 <> 0) Then
errorname:
        MsgBox ("Function BootStrap.Is_Module_Loaded Was not passed a Name of Module")
        Stop
    End If
    If (0 <> 0) Then
notfound:
        MsgBox ("MODULE: " & name & " is not installed please add")
        Stop
    End If
End Function

'----- Exists For Collections
Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant: Set var = Nothing
  Dim errNumber As Long
  InCollection = False
  
  Err.Clear
  On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

