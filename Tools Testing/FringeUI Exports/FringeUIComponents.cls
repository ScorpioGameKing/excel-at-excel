VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- Package and Struct Vars
Private tabs As Collection
Public UIPackage As Collection
'-----
'----- Init and Builds. Build Ensures the package is structured correctly
Private Sub Class_Initialize()
    If tabs Is Nothing Then Set tabs = New Collection
    If UIPackage Is Nothing Then Set UIPackage = New Collection
End Sub

Public Sub BuildPackage()
    Set UIPackage = tabs
End Sub
'-----
'----- TAB STRUCT: An ID to search, An ID for build, the XML to inject, and a Collection to hold Groups
Private Sub AddTabStruct(id As String, label As String, insertBeforeQ As String)
    Dim tab_struct As Collection: Set tab_struct = New Collection
    Dim tab_contents As Collection: Set tab_contents = New Collection
    Dim header As String
    
    header = "<mso:tab id='" & id & _
    "' label='" & label & _
    "' insertBeforeQ='" & insertBeforeQ & "'>" & vbNewLine
    header = Replace(header, "'", """")
    
    tab_struct.Add "TAB": tab_struct.Add header: tab_struct.Add tab_contents, "CONTENTS"
    If Not InCollection(tabs, id) Then
        tabs.Add tab_struct, id
    Else
        tabs.Remove (id)
        tabs.Add tab_struct, id
    End If
End Sub

Public Sub AddTab(id As String, label As String, insertBeforeQ As String)
    AddTabStruct id, label, insertBeforeQ
End Sub
'-----
'----- GROUP STRUCT: An ID to search, An ID for build, the XML to inject, and a Collection to hold Buttons, Menus, ETC
Private Sub AddGroupStruct(tId As String, id As String, label As String, autoScale As String)
    Dim group_struct As Collection: Set group_struct = New Collection
    Dim group_contents As Collection: Set group_contents = New Collection
    Dim header As String
    
    header = "<mso:group id='" & id & _
    "' label='" & label & _
    "' autoScale='" & autoScale & "'>" & vbNewLine
    header = Replace(header, "'", """")
    
    group_struct.Add "GROUP": group_struct.Add header: group_struct.Add group_contents, "CONTENTS"
    If Not InCollection(tabs(tId).Item("CONTENTS"), id) Then
        tabs(tId).Item("CONTENTS").Add group_struct, id
    Else
        tabs(tId).Item("CONTENTS").Remove (id)
        tabs(tId).Item("CONTENTS").Add group_struct, id
    End If
End Sub

Public Sub AddGroup(tId As String, id As String, label As String, autoScale As String)
    AddGroupStruct tId, id, label, autoScale
End Sub
'-----
'----- BUTTON STRUCT: A TAG for build, XML to Inject , Tab ID to sort
Private Sub AddButtonXML(tId As String, gId As String, id As String, label As String, image As String, callback As String)
    Dim button_struct As Collection: Set button_struct = New Collection
    
    bt = "<mso:button id='" & id & _
    "' label='" & label & _
    "' imageMso='" & image & _
    "' onAction='" & callback & "'/>" & vbNewLine
    bt = Replace(bt, "'", """")
    
    button_struct.Add "BUTTON": button_struct.Add bt
    If Not InCollection(tabs.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS"), id) Then
        tabs.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Add button_struct, id
    Else
        tabs.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Remove (id)
        tabs.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Add button_struct, id
    End If
End Sub

Public Sub AddButton(tId As String, gId As String, id As String, label As String, image As String, callback As String)
    AddButtonXML tId, gId, id, label, image, callback
End Sub
'-----
'----- Exists For Collections
'----- Credit: https://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

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
'-----



