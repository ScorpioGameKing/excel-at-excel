VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GenerateButton_Click()
    '-------------------
    'Initialization Data
    '-------------------
    
    'Sheet IDs
    Dim tool_id As Worksheet
    Dim template_id As Worksheet
    Dim db_id As Worksheet
    Dim place_after_id As Worksheet
        
    'Variables
    Dim week_start As Range
    Dim week_length As Range
    Dim date_range As Range
    Dim active_ref As String
    
    'DB Variables
    Dim last_date_range As Range
    Dim last_date_added As String
    Dim date_db As ListObject
    
    'Set IDs
    Set tool_id = Worksheets("Tool Page")
    Set template_id = Worksheets("Template Sheet")
    
    'Set Variables
    Set week_start = Range("B1")
    Set week_length = Range("B2")
    Set date_range = Range("B4:B7")
    
    'Set DB Variables
    Set last_date_range = Range("B10")
    last_date_added = last_date_range.Value
    Set date_db = ListObjects("date_db")
    
    'Debug Prints
    Debug.Print "Week Starting On: " & week_start.Value
    Debug.Print "Week lasts: " & week_length.Value & " Days"
    Debug.Print "Tool Page ID: " & tool_id.Name
    Debug.Print "Template Page ID: " & template_id.Name
    
    'Validation Check
    Dim validation As Boolean
    validation = DateInDB(Replace(week_start.Value, "/", "-"), date_db)
    
    'If the day we're tyring to start on hasn't been made
    If Not validation Then
    
        'Decide the placement ID
        If last_date_added = "" Then
            Set place_after_id = Worksheets("Tool Page")
        Else
            Debug.Print (last_date_added)
            Set place_after_id = Worksheets(Replace(last_date_added, "/", "-"))
        End If
        
        'Iterate through the Dates to add. Make a DB entry and new Sheet for each. Name Sheet and Header appropriately
        For Each cell In date_range
            Dim db_entry As ListRow
            
            'If this is the first date we have no active reference
            If active_ref = "" Then
                Debug.Print "Creating sheet for: " & Replace(cell.Value, "/", "-")
                template_id.Copy After:=place_after_id
                active_ref = Replace(cell.Value, "/", "-")
                ActiveSheet.Name = Replace(cell.Value, "/", "-")
                ActiveSheet.Range("A1").Value = active_ref
            'Otherwise we do and adjust the order based off it
            Else
                Debug.Print "Creating sheet for: " & Replace(cell.Value, "/", "-")
                template_id.Copy After:=Worksheets(active_ref)
                active_ref = Replace(cell.Value, "/", "-")
                ActiveSheet.Name = Replace(cell.Value, "/", "-")
                ActiveSheet.Range("A1").Value = active_ref
            End If
            
            'Add the sheet date to the DB
            Set db_entry = date_db.ListRows.Add(AlwaysInsert:=True)
            db_entry.Range.Value = active_ref
        Next cell
        
        'After everything save a carry over reference to offset future generations
        tool_id.Activate
        last_date_added = active_ref
        last_date_range.Value = last_date_added
    
    'The date already exists and would cause it and others to be overwritten. MANUAL OVERRIDE REQUIRED TO REMOVE DATES
    Else
        MsgBox "Invalid Start Date. Date and following dates are already in the database"
    End If
End Sub

'Validation Check. Iter the DB for starting date. Plan to add offset from last_added_date for speed increase before a full index attempt
Private Function DateInDB(date_val As String, db As ListObject) As Boolean
    For Each Row In db.ListRows
        Debug.Print "Expected Check: " & Replace(Row.Range.Value, "/", "-")
        Debug.Print "Check For: " & date_val
        If date_val = Replace(Row.Range.Value, "/", "-") Then
            'If True then Exit cause speed is key in the long run here
            DateInDB = True
            Exit Function
        End If
    Next Row
    DateInDB = False
End Function
