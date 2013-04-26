' ------------------------------------------------------------
' Name: CleanupUtils.vbs
' Description: Contains utility functions related to cleanup
' ------------------------------------------------------------
Attribute VB_Name = "GeneralUtils"

' ------------------------------------------------------------------------------------------------------------------------

' Name: Expiration30Days
' Description: Marks the selected emails to expire in 30 days
Sub Expiration30Days()

    Dim objApp As Outlook.Application
        
    Set objApp = CreateObject("Outlook.Application")
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
        Dim objOpenItem As Object
        
        If Application.ActiveExplorer.Selection.count = 0 Then
            Exit Sub
        End If
    
        For Each objItem In Application.ActiveExplorer.Selection
            If objItem.Class = olMail Then
                Set objOpenItem = objItem
                Call SetExpirationFlag(objOpenItem, 30)
            End If
        Next
    Case "Inspector"
        Set objOpenItem = objApp.ActiveInspector.CurrentItem
        Call SetExpirationFlag(objOpenItem, 30)
    Case Else
            ' anything else will result in an error, which is
            ' why we have the error handler above
    End Select
      
    Set objOpenItem = Nothing
        
End Sub

' Name: SetExpirationFlag
' Description: Sets the expiration date of the specified mail item to X days.
' Arguments:
'			objItem - The item to set
'			expiration - The number of days for expiration. Negative numbers will cause the item to be expired immediatly
Sub SetExpirationFlag(objItem As Object, expiration As Integer)
    With objItem
        If .Class = olMail Then
            .ExpiryTime = Date + expiration
            .Save
        End If
    End With
End Sub

' Name: NeverExpire
' Description: Marks the selected emails to never expire. 
'			   Has no effect if the mail has no expiration flag.
Sub NeverExpire()

    Dim objApp As Outlook.Application
        
    Set objApp = CreateObject("Outlook.Application")
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
        Dim objOpenItem As Object
        
        If Application.ActiveExplorer.Selection.count = 0 Then
            Exit Sub
        End If
    
        For Each objItem In Application.ActiveExplorer.Selection
            If objItem.Class = olMail Then
                Set objOpenItem = objItem
                Call UnSetExpirationFlag(objOpenItem)
            End If
        Next
    Case "Inspector"
        Set objOpenItem = objApp.ActiveInspector.CurrentItem
        Call UnSetExpirationFlag(objOpenItem)
    Case Else
            ' anything else will result in an error, which is
            ' why we have the error handler above
    End Select
      
    Set objOpenItem = Nothing
        
End Sub

' Name: UnSetExpirationFlag
' Description: Un-Sets the expiration date of the specified mail item.
' Arguments:
'			objItem - The item to set
Sub UnSetExpirationFlag(objItem As Object)
    With objItem
        If .Class = olMail Then
            .ExpiryTime = DateTime.DateSerial(4501, 1, 1) ' VBA Stuff - this date signifies that the item never expires.
            .Save
        End If
    End With
End Sub