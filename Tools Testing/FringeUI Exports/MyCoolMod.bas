Attribute VB_Name = "MyCoolMod"
Private FringeUI As Object
Private uiPackage As Object

Public Sub InitUI(Optional multiLoader As Variant)
    If FringeUI Is Nothing Then Set FringeUI = New FringeUIManager
    If uiPackage Is Nothing Then Set uiPackage = New FringeUIPackage
    
    uiPackage.AddTab "MyCustomTab", "My New Tab", "mso:TabFormat"
    uiPackage.AddGroup "MyCustomTab", "MyNewGroup", "My New Group", "true"
    uiPackage.AddButton "MyCustomTab", "MyNewGroup", "HelloWorld", "Hello World!", "PanningHand", "MyCoolMod.HelloWorld"
    
    If multiLoader Is Nothing Then
        FringeUI.BuildFringeUI uiPackage.uiPackage, True
    Else
        multiLoader.AddUIPackage uiPackage, "MyCoolUI"
    End If
End Sub

Public Sub HelloWorld()
    MsgBox "Hello World!", vbOKOnly, "Hi There!"
End Sub

Public Sub HelloOtherGroup()
    MsgBox "Hello Other Group!", vbOKOnly, "Hi There Again!"
End Sub
