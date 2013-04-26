' ------------------------------------------------------------
' Name: GeneralUtils.vbs
' Description: Contains general utility functions related 
' ------------------------------------------------------------
Attribute VB_Name = "GeneralUtils"

' ------------------------------------------------------------------------------------------------------------------------

' Name: GetCurrentItem
' Description: Gets the currently selected mail item. and returns it as an Object.
' Returns: The selected mail item. If the current view is the "Explorer" and there is more than one item selected - ONLY THE FIRST is returned.
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
        
    Set objApp = CreateObject("Outlook.Application")
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        Case Else
            ' anything else will result in an error, which is
            ' why we have the error handler above
    End Select
    
    Set objApp = Nothing
End Function

' ------------------------------------------------------------------------------------------------------------------------
