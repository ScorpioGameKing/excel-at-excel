VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- FRINGEUI: VBA Wasn't Meant For This
' Author: ScorpioGameKing
' Version: 0.1.0
'
' The one stop toolkit for injecting buttons into the Office Ribbon.
' Want your custom module to feel just that little  bit more complete?
' Try FringeUI and in 1 line add your own custom buttons that call *your*
' module's Subs and Functions. Want even more? Add custom tabs for each
' of you modules and include help buttons, custom commands, custom
' autocompletes, as long as you can write it, FringeUI can add it.
'
' SUPPORTED:
'   TABS id (String), label (String), position (insertBeforeQ:mso (id))
'   GROUPS tId (String), id (String), label (String), scaling (Boolean)
'   BUTTONS tId (String), gId (String), id (String), label (String), image (imageMso), callback (Module.Sub)
'
' PLANNED:
'   MENU
'   More... idk but more
'-----

'----- XML INJECTION PHASE
Sub BuildFringeUI(package As Collection, Optional DEBUGMODE As Boolean)

    ClearFringeUI DEBUGMODE

    '----- Prep Variables
    Dim hFile As Long: hFile = FreeFile
    Dim path As String: path = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
    Dim fileName As String: fileName = "Excel.officeUI"
    Dim ribbonXML As String
    
    '----- Nothing to do
    ribbonXML = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine & _
    "<mso:ribbon startFromScratch='false'>" & vbNewLine & _
    "<mso:tabs>" & vbNewLine
    
    '----- Start Tab Injections
    For Each tab_check In package
    
        '----- Validate
        If tab_check.Item(1) = "TAB" Then
        
            '----- Insert Header, Iterate
            ribbonXML = ribbonXML & tab_check.Item(2)
            For Each tab_group_check In tab_check.Item(3)
            
                '----- Validate
                If tab_group_check.Item(1) = "GROUP" Then
                
                    '----- Insert Header, Iterate
                    ribbonXML = ribbonXML & tab_group_check.Item(2)
                    For Each tab_obj_check In tab_group_check.Item(3)
                    
                        '----- Validate
                        If tab_obj_check.Item(1) = "BUTTON" Then
                        
                            '----- Insert Button
                            ribbonXML = ribbonXML & tab_obj_check.Item(2)
                        End If
                    Next
                    
                    '----- Insert Group Exit
                    ribbonXML = ribbonXML & "</mso:group>" & vbNewLine
                End If
            Next
            
            '----- Insert Tab Exit
            ribbonXML = ribbonXML & "</mso:tab>" & vbNewLine
        End If
    Next
    
    '----- Close Out Misc Tags
    ribbonXML = ribbonXML & "</mso:tabs>" & vbNewLine & _
    "</mso:ribbon>" & vbNewLine & _
    "</mso:customUI>"
    
    '----- Format to proper XML
    ribbonXML = Replace(ribbonXML, "'", """")
    
    '----- Inject the custom ribbon
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile
    
    '----- Debug Display
    If DEBUGMODE Then MsgBox ribbonXML
    If DEBUGMODE Then Debug.Print ribbonXML

End Sub
'-----

'----- CLEAN UP AFTER OURSELVES
Sub ClearFringeUI(Optional DEBUGMODE As Boolean)

    Dim hFile As Long: hFile = FreeFile
    Dim path As String: path = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
    Dim fileName As String: fileName = "Excel.officeUI"
    Dim ribbonXML As String
    
    ribbonXML = Replace("<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & _
    "<mso:ribbon></mso:ribbon></mso:customUI>" _
    , "'", """")
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile
    
    '----- Debug Display
    If DEBUGMODE Then MsgBox ribbonXML
    If DEBUGMODE Then Debug.Print ribbonXML

End Sub
'-----




