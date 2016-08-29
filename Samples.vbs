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

Sub ArchiveSelectedMessages()
    'On Error Resume Next
    
    Dim objItem As Outlook.MailItem
    
    If Application.ActiveExplorer.Selection.count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    Dim yearStart As Date
    yearStart = DateSerial(2010, 1, 1)
    For Each objItem In Application.ActiveExplorer.Selection
        'If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                Dim stroeId As String
                If objItem.SentOn >= yearStart Then
                    stroeId = "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000000043003A005C006D00610069006C005C0057006F0072006B0020002D00200032003000310030002E007000730074000000"
                Else
                    stroeId = "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000000043003A005C006D00610069006C005C0057006F0072006B0020002D00200032003000300039002E007000730074000000"
                End If
                
                ArchiveMessage objItem, stroeId
                
            End If
        'End If
    Next

    Set objItem = Nothing
    
End Sub

Sub ArchiveMessage(objItem As Outlook.MailItem, stroeId As String)
    Dim targetFolder As Outlook.MAPIFolder
    Set targetFolder = GetArchiveFolder(objItem.Parent, stroeId)
    
    objItem.UnRead = False
    objItem.Move targetFolder
End Sub

Function GetArchiveFolder(sourceFolder As Outlook.MAPIFolder, stroeId As String) As Outlook.MAPIFolder
    Dim objGrandParent As Object
    Set objGrandParent = sourceFolder.Parent.Parent
    Dim objParent As Object
    Set objParent = sourceFolder.Parent
    Dim objTargetFolder As Outlook.MAPIFolder
    
    If objGrandParent.Class = olNamespace Then
        Dim objPST As Outlook.MAPIFolder
        Set objPST = objGrandParent.GetFolderFromID(stroeId)
        Set objTargetFolder = FindOrCreate(objPST, sourceFolder.Name)
        Set GetArchiveFolder = objTargetFolder
    Else
        Dim parentFolder As Outlook.MAPIFolder
        Set parentFolder = GetArchiveFolder(sourceFolder.Parent, stroeId)
        Set objTargetFolder = FindOrCreate(parentFolder, sourceFolder.Name)
        Set GetArchiveFolder = objTargetFolder
    End If
    
End Function

Function FindOrCreate(parentFolder As Outlook.MAPIFolder, folderName As String) As Outlook.MAPIFolder
    Dim count As Integer
    count = parentFolder.folders.count
    
    Dim index As Integer
    index = 1
    
    Dim folder As Outlook.MAPIFolder
    
    Do While index <= count
        Set folder = parentFolder.folders.Item(index)
        If folder.Name = folderName Then
            Set FindOrCreate = folder
        End If
        index = index + 1
    Loop
    
    If FindOrCreate Is Nothing Then
        Set FindOrCreate = parentFolder.folders.Add(folderName)
    End If
    
End Function

Sub ArchiveSelectedMessagesTo_RMF_Folder()
    
    Dim objNS As Outlook.NameSpace
    Set objNS = Application.GetNamespace("MAPI")
    
    Dim objPST As Outlook.MAPIFolder
    Set objPST = objNS.GetFolderFromID("0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000000043003A005C006D00610069006C005C0057006F0072006B0020002D00200032003000310030002E007000730074000000")
    
    Dim objRMFFolder As Outlook.MAPIFolder
    Set objRMFFolder = objPST.folders("_RMF")
        
    Dim objItem As Outlook.MailItem

    'Assume this is a mail folder
    If objRMFFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
    End If

    If Application.ActiveExplorer.Selection.count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    For Each objItem In Application.ActiveExplorer.Selection
        If objRMFFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                objItem.UnRead = False
                objItem.Move objRMFFolder
            End If
        End If
    Next

    Set objItem = Nothing
    Set objRMFFolder = Nothing
    Set objPST = Nothing
    Set objNS = Nothing
    
End Sub


Sub ArchiveSelectedMessagesTo_ABI_Folder()
   
On Error Resume Next

    Dim objNS As Outlook.NameSpace
    Set objNS = Application.GetNamespace("MAPI")
    
    Dim objPST As Outlook.MAPIFolder
    Set objPST = objNS.GetFolderFromID("0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000000043003A005C006D00610069006C005C0057006F0072006B0020002D00200032003000310030002E007000730074000000")
    
    Dim objRMFFolder As Outlook.MAPIFolder
    Set objRMFFolder = objPST.folders("_RMF")
    
    Dim objABIFolder As Outlook.MAPIFolder
    Set objABIFolder = objRMFFolder.folders("_ABI")
        
    Dim objItem As Outlook.MailItem

    'Assume this is a mail folder
    If objABIFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
    End If

    If Application.ActiveExplorer.Selection.count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    For Each objItem In Application.ActiveExplorer.Selection
        If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                objItem.UnRead = False
                objItem.Move objABIFolder
            End If
        End If
    Next

    Set objItem = Nothing
    Set objABIFolder = Nothing
    Set objRMFFolder = Nothing
    Set objPST = Nothing
    Set objNS = Nothing

End Sub

Sub MoveSelectedMessagesTo_Inbox()

On Error Resume Next

    Dim objFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace, objItem As Outlook.MailItem
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

Sub MoveSelectedMessagesTo_ABI_Folder()

On Error Resume Next

    Dim objFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace, objItem As Outlook.MailItem
    Set objNS = Application.GetNamespace("MAPI")
       
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
    Set objRMFFolder = objInbox.Parent.folders("_RMF")
    Set objFolder = objRMFFolder.folders("_ABI")

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

Sub MoveSelectedMessagesTo_RMF_Folder()

On Error Resume Next

    Dim objFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace, objItem As Outlook.MailItem
    Set objNS = Application.GetNamespace("MAPI")
       
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
    Set objFolder = objInbox.Parent.folders("_RMF")

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

Sub MoveSelectedMessagesTo_Galileo_Folder()

On Error Resume Next

    Dim objFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace, objItem As Outlook.MailItem
    Set objNS = Application.GetNamespace("MAPI")
       
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)
    Set objFolder = objInbox.folders("Galileo")

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
Sub CleanUpConvesations()
    Set objTempItem = Application.ActiveExplorer.Selection
    Set objFolder = objTempItem.Parent.CurrentFolder
    Dim objOpenItem As Object
    Dim currentMailItem As MailItem
    
    rootConversations As New Collection
        
    For Each objItem In objFolder.Items
        If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                Set currentMailItem = objItem
                
                Set currentThreadItem = New ThreadItem
                currentThreadItem.AddThreadItem currentMailItem.ConversationIndex, currentMailItem, New Collection
                rootConversations.Add currentThreadItem
            End If
        End If
    Next
                
End Sub
Sub CleanUpDir()
    Set objTempItem = Application.ActiveExplorer.Selection
    Set objFolder = objTempItem.Parent.CurrentFolder
    Dim objOpenItem As Object
    
    For Each objItem In objFolder.Items
        If objFolder.DefaultItemType = olMailItem Then
            If objItem.Class = olMail Then
                
                If StringStartsWith(objItem.Subject, "Nepal Defects Daily Report") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
                If StringStartsWith(objItem.Subject, "Maintenance Daily Report") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
               If StringStartsWith(objItem.Subject, "New 750 daily storages are available") Or StringStartsWith(objItem.Subject, "RE: New 750 daily storages are available") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
                If StringStartsWith(objItem.Subject, "New 700 daily storages are available") Or StringStartsWith(objItem.Subject, "RE: New 700 daily storages are available") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
                
                If StringStartsWith(objItem.Subject, "Night Build") Or StringStartsWith(objItem.Subject, "RE: Night Build") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
                If StringStartsWith(objItem.Subject, "PCI-ABP, Night Build Report") Or StringStartsWith(objItem.Subject, "RE: PCI-ABP, Night Build Report") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
                If StringStartsWith(objItem.Subject, "New V800 packages are available") Then
                    Set objOpenItem = objItem
                    Call Clean(objOpenItem)
                End If
                
            End If
        End If
    Next

End Sub

Sub Clean(objMailItem As Object)
    Call SetExpirationFlag(objMailItem, 30)
    objMailItem.UnRead = False
End Sub
Sub Blue_Code_Highlight()
    Dim msg As Outlook.MailItem
    Dim insp As Outlook.Inspector
    
    Set insp = Application.ActiveInspector
    If insp.CurrentItem.Class = olMail Then
        Set msg = insp.CurrentItem
        Selection.Font.Color = wdColorRed
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 10
        'olEditorWord
        If insp.EditorType = olEditorHTML Then
            Set hed = msg.GetInspector.HTMLEditor
            Set rng = hed.Selection.createRange
            rng.pasteHTML "<font style='color: blue; font-family:Times New Roman; font-size: 10pt;'>" & rng.Text & "</font><br/>"
        End If
    End If
    Set insp = Nothing
    Set rng = Nothing
    Set hed = Nothing
    Set msg = Nothing
End Sub