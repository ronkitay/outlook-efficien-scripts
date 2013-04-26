' ------------------------------------------------------------
' Name: Samples.vbs
' Description: Contains samples of code for doing things with VB in Outlook
' ------------------------------------------------------------
Attribute VB_Name = "Samples"

' ------------------------------------------------------------------------------------------------------------------------

Sub DoSomethingToAllMailItemsInCurrentDirectory()
    Set objTempItem = Application.ActiveExplorer.Selection
    Set objFolder = objTempItem.Parent.CurrentFolder
    Dim objOpenItem As Object
	Dim currentMailItem As MailItem
    
    For Each objItem In objFolder.Items
        If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                
                Set currentMailItem = objItem
				
                If StringStartsWith(currentMailItem.Subject, "[SPAM]") Then
                    Call Clean(currentMailItem)
                End If
                
            End If
        End If
    Next

End Sub

Sub Clean(objMailItem As Object)
    Call SetExpirationFlag(objMailItem, 30)
    objMailItem.UnRead = False
End Sub

Sub MoveSelectedMessagesTo_Inbox()

On Error Resume Next

    Dim objFolder As Outlook.MAPIFolder
	Dim objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace
	Dim objItem As Outlook.MailItem
    Set objNS = Application.GetNamespace("MAPI")
       
    Set objFolder = objNS.GetDefaultFolder(olFolderInbox)

    'Assume this is a mail folder
    If objFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
    End If

    If Application.ActiveExplorer.Selection.count = 0 Then

        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    For Each objItem In Application.ActiveExplorer.Selection
        If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                objItem.Move objFolder
            End If
        End If
    Next

    Set objItem = Nothing
    Set objFolder = Nothing
    Set objInbox = Nothing
    Set objNS = Nothing

End Sub

Function FindOrCreateFolder(parentFolder As Outlook.MAPIFolder, folderName As String) As Outlook.MAPIFolder
    Dim count As Integer
    count = parentFolder.folders.count
    
    Dim index As Integer
    index = 1
    
    Dim folder As Outlook.MAPIFolder
    
    Do While index <= count
        Set folder = parentFolder.folders.Item(index)
        If folder.Name = folderName Then
            Set FindOrCreateFolder = folder
        End If
        index = index + 1
    Loop
    
    If FindOrCreateFolder Is Nothing Then
        Set FindOrCreateFolder = parentFolder.folders.Add(folderName)
    End If
    
End Function