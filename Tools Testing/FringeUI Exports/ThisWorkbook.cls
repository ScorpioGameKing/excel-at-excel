VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private UIManager As New FringeUIManager
Private MyCustomUI As New FringeUIPackage
Private multiLoader As Object

Private Sub Workbook_Activate()
    If RealityCheck.IsModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsModuleLoaded("FringeUIManager") * RealityCheck.IsModuleLoaded("FringeUIPackage") Then
        ' Multiple UI Packages
        Set multiLoader = New FringeUIMultiLoader
        PPTLinker.Init "PPTLinking", "PPTLinking", ".pptx", multiLoader
        MyCoolMod.InitUI multiLoader
        multiLoader.BuildMultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    ElseIf RealityCheck.IsModuleLoaded("FringeUIManager") * RealityCheck.IsModuleLoaded("FringeUIPackage") Then
        ' Single UI Package
        PPTLinker.Init "PPTLinking", "PPTLinking", ".pptx"
        multiLoader.BuildMultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    Else
        ' Default Launch
    End If
End Sub

Private Sub Workbook_Deactivate()
    UIManager.ClearFringeUI
End Sub
