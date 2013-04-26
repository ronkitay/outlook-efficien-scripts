' ------------------------------------------------------------
' Name: SearchUtils.vbs
' Description: Contains all functions related to Searching
' ------------------------------------------------------------
Attribute VB_Name = "SearchUtils"

' ------------------------------------------------------------------------------------------------------------------------

' Name: FindSelectedMessages
' Description: Finds all emails with subject like the current selected email. 
Sub FindSelectedMessages()
    'On Error Resume Next
    
    Dim objItem As Outlook.MailItem
    
    If Application.ActiveExplorer.Selection.count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If
    
    If Application.ActiveExplorer.Selection.count > 1 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If
    
    Dim msg As Outlook.MailItem
    Set msg = Application.ActiveExplorer.Selection.Item(1)
      
    Call Application.ActiveExplorer.Search("""" & msg.Subject & """", olSearchScopeAllFolders)
    
End Sub