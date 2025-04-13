VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIMultiLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private UIPackages As New Collection
Public MultiUIPackage As New Collection

Public Sub AddUIPackage(uiPackage As FringeUIPackage, id As String)
    If Not RealityCheck.InCollection(UIPackages, id) Then UIPackages.Add uiPackage, id
End Sub

Public Sub RemoveUIPackage(uiPackage As FringeUIPackage, id As String)
    If RealityCheck.InCollection(UIPackages, id) Then UIPackages.Remove uiPackage, id
End Sub

Public Sub BuildMultiUIPackage()
    For Each sub_pack_tab In UIPackages
        MultiUIPackage.Add sub_pack_tab.uiPackage.Item(1)
    Next
End Sub

