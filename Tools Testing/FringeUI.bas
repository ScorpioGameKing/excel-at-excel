'------------------------------------------ FRINGEUI: VBA Wasn't Meant For This ------------------------------------------
' Author: ScorpioGameKing
' Version: 0.0.3
'
' The one stop toolkit for injecting buttons into the Office Ribbon. Want your custom module to feel just that little
' bit more complete? Try FringeUI and in 1 line add your own custom buttons that call *your* module's Subs and Functions.
' Want even more? Add custom tabs for each of you modules and include help buttons, custom commands, custom autocompletes,
' as long as you can write it, FringeUI can add it.
'
' SUPPORTED:
'   BUTTONS id, label, image, callback *MUST POINT TO MODULE SUB/FUNC"
'
' PLANNED:
'   TAB id, label, position
'   More... idk but more
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------- THE INJECTION COLLECTION ---------------------------------------------------
Private buttons As New Collection
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------- XML INJECTION PHASE --------------------------------------------------------
Sub LoadCustRibbon()

'----- Prep Variables
Dim hFile As Long: hFile = FreeFile
Dim path As String: path = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
Dim fileName As String: fileName = "Excel.officeUI"
Dim ribbonXML As String

'----- Nothing to do
ribbonXML = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
ribbonXML = ribbonXML & "<mso:ribbon startFromScratch='false'>" & vbNewLine
ribbonXML = ribbonXML & "<mso:tabs>" & vbNewLine

'----- Start Tab Injections, FringeUI only if no Custom Tab is Provided
ribbonXML = ribbonXML & "<mso:tab id='FringeUI' label='FringeUI' insertBeforeQ='mso:TabFormat'>" & vbNewLine

'----- Create Button Groups, Default Group is FringeUIGroup if none Provided
ribbonXML = ribbonXML & "<mso:group id='FringeUIGroup' label='FringeUI' autoScale='true'>" & vbNewLine

'----- Inject Custom Buttons to Group if provided, Default is Hello World Button
If buttons.Count = 0 Then
    ribbonXML = ribbonXML & "<mso:button id='helloWorldTest' label='Hello World' imageMso='PanningHand' onAction='FringeUI.HelloWorld'/>" & vbNewLine
Else
    For Each b In buttons
        ribbonXML = ribbonXML & b & vbNewLine
    Next
End If

'----- Close out Tags
ribbonXML = ribbonXML & "</mso:group>" & vbNewLine
ribbonXML = ribbonXML & "</mso:tab>" & vbNewLine
ribbonXML = ribbonXML & "</mso:tabs>" & vbNewLine
ribbonXML = ribbonXML & "</mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML & "</mso:customUI>"

'----- Format to proper XML
ribbonXML = Replace(ribbonXML, "'", """")

'----- Inject the custom ribbon
Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

'----- Debug Display
MsgBox ribbonXML

End Sub
'-------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------- CLEAN UP AFTER OURSELVES ------------------------------------------------
Sub ClearCustRibbon()

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

End Sub
'-------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------- DEBUG -----------------------------------------------------------
Public Sub HelloWorld()
    MsgBox ("Hello World!")
End Sub
'-------------------------------------------------------------------------------------------------------------------------

'--------------------------------------------------- BUTTON INJECTION ----------------------------------------------------
Private Sub AddButtonXML(id As String, label As String, image As String, callback As String)
    Dim bt As String
    bt = "<mso:button id='" & id & _
    "' label='" & label & _
    "' imageMso='" & image & _
    "' onAction='" & callback & "'/>"
    bt = Replace(bt, "'", """")
    buttons.Add bt
End Sub

Public Sub AddButton(id As String, label As String, image As String, callback As String)
    AddButtonXML id, label, image, callback
End Sub
'-------------------------------------------------------------------------------------------------------------------------

