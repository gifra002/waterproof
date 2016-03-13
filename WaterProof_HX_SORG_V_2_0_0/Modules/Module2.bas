Attribute VB_Name = "Module2"
Public Unlocked As Boolean
Public Const APP_NAME = "WaterProof_HX"
Public Const SECTION_NAME = "Gifra"
Public Const KEY_NAME = "Deltagifra"
Public Sub CheckLockedStatus(temp)
temp = GetSetting(APP_NAME, SECTION_NAME, KEY_NAME, "locked")
    If temp = "unlocked" Then
        Unlocked = True
    Else
        Unlocked = False
    End If
End Sub


