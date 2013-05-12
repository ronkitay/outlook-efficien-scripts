' ------------------------------------------------------------
' Name: MailEditing.vbs
' Description: Contains all functions related to Mail Editing
' ------------------------------------------------------------
Attribute VB_Name = "MailEditing"

' ------------------------------------------------------------------------------------------------------------------------

' Name: EmphesizeSelectedText
' Description: Turns the selected text BOLD and changes its color to the specified color.
'			   Works on a single word even if the cursor is on it and it is not selected.
' Arguments:
'			color - The new color for the text. Accepts arguments such as "RGB(0,255,0)"
' Limitations: Works only when using the "Word" editor
Sub EmphesizeSelectedText(color As Long)
    Dim msg As Outlook.MailItem
    Dim insp As Outlook.Inspector

    Set insp = Application.ActiveInspector
    If insp.CurrentItem.Class = olMail Then
        Set msg = insp.CurrentItem
        If insp.EditorType = olEditorWord Then
            
            Set document = msg.GetInspector.WordEditor
            Set rng = document.Application.Selection
            
            With rng.Font
                .Bold = True
                .color = color
            End With
        End If
    End If
    Set insp = Nothing
    Set rng = Nothing
    Set hed = Nothing
    Set msg = Nothing
End Sub

' Name: EmphesizeSelectedTextAsBlue
' Description: Turns the selected text BOLD and changes its color to be Blue - RGB(0, 112, 202)
Sub EmphesizeSelectedTextAsBlue()
    EmphesizeSelectedText (RGB(0, 112, 202))
End Sub

' Name: EmphesizeSelectedTextAsRed
' Description: Turns the selected text BOLD and changes its color to be Red - RGB(255, 0, 0)
Sub EmphesizeSelectedTextAsRed()
    EmphesizeSelectedText (RGB(255, 0, 0))
End Sub

' Name: EmphesizeSelectedTextAsGreen
' Description: Turns the selected text BOLD and changes its color to be Blue - Green(0, 176, 80)
Sub EmphesizeSelectedTextAsGreen()
    EmphesizeSelectedText (RGB(0, 176, 80))
End Sub

' Name: EmphesizeSelectedTextAsOrange
' Description: Turns the selected text BOLD and changes its color to be Orange - RGB(255, 128, 0)
Sub EmphesizeSelectedTextAsOrange()
    EmphesizeSelectedText (RGB(255, 128, 0))
End Sub

' Name: EmphesizeSelectedTextAsPurple
' Description: Turns the selected text BOLD and changes its color to be Purple - RGB(112, 48, 177)
Sub EmphesizeSelectedTextAsPurple()
    EmphesizeSelectedText (RGB(112, 48, 177))
End Sub

' ------------------------------------------------------------------------------------------------------------------------

' Name: HighlightSelectedText
' Description: Changes the background color of the selected text to the specified color.
' Arguments:
'			color - The new color for the text. Accepts arguments such as "RGB(0,255,0)"
' Limitations: Works only when using the "Word" editor
Sub HighlightSelectedText(color As Long)
    Dim msg As Outlook.MailItem
    Dim insp As Outlook.Inspector

    Set insp = Application.ActiveInspector
    If insp.CurrentItem.Class = olMail Then
        Set msg = insp.CurrentItem
        If insp.EditorType = olEditorWord Then
            
            Set document = msg.GetInspector.WordEditor
            Set rng = document.Application.Selection
            
            With rng.Font
                .Shading.BackgroundPatternColor = color
            End With
        End If
    End If
    Set insp = Nothing
    Set rng = Nothing
    Set hed = Nothing
    Set msg = Nothing
End Sub

' Name: HighlightSelectedTextAsGrey
' Description: Changes the background color of the selected text to be Grey - RGB(192, 192, 192)
Sub HighlightSelectedTextAsGrey()
    HighlightSelectedText (RGB(192, 192, 192))
End Sub

' Name: HighlightSelectedTextAsYellow
' Description: Changes the background color of the selected text to be Yellow - RGB(255, 255, 0)
Sub HighlightSelectedTextAsYellow()
    HighlightSelectedText (RGB(255, 255, 0))
End Sub

' Name: HighlightSelectedTextAsRed
' Description: Changes the background color of the selected text to be Red - RGB(255, 0, 0)
Sub HighlightSelectedTextAsRed()
    HighlightSelectedText (RGB(255, 0, 0))
End Sub

' ------------------------------------------------------------------------------------------------------------------------

' Name: AddCodeTextBox
' Description: Adds a floating text box which is suitable for code snipets.
'			   Creates a Black text box, with White text color and a dark purple border (4px)
' Limitations: Works only when using the "Word" editor.
'			   The text box needs to be manually "inlined wtih text" for it to be comfortable to use
Sub AddCodeTextBox()
    Dim msg As Outlook.MailItem
    Dim insp As Outlook.Inspector
   
    Set insp = Application.ActiveInspector
    If insp.CurrentItem.Class = olMail Then
        Set msg = insp.CurrentItem
        If insp.EditorType = olEditorWord Then
            
            Dim document As Object
            Set document = msg.GetInspector.WordEditor
            
            Dim rng As Object
            Set rng = document.Application.Selection
            
            Dim shape As Object
            Set shape = document.Shapes.AddTextbox(msoTextOrientationHorizontal, 12, 12, 600#, 150#)
            
            shape.Select
            shape.Fill.BackColor = RGB(255, 255, 255)
            shape.Fill.ForeColor = RGB(0, 0, 0)
            shape.Fill.Solid
            shape.Line.Weight = 4
            shape.Line.DashStyle = msoLineSolid
            shape.Line.Style = msoLineSingle
            shape.Line.Transparency = 0#
            shape.Line.ForeColor = RGB(77, 3, 47)
                        
        End If
    End If
    Set insp = Nothing
    Set msg = Nothing
End Sub

' ------------------------------------------------------------------------------------------------------------------------

' Name: SetDelayedDeliveryToNextWorkingDay
' Description: Marks the current email (edit mode) to be sent during the next working day.
'			   If the day is Thursday or beyond - skips to Sunday.
'			   Email sending time is always set to 7:<current minutes>. E.g., if the time is now 19:32 and the macro is executed, the 
'			   email will be delivered at 7:32
Sub SetDelayedDeliveryToNextWorkingDay()
    Dim msg As Outlook.MailItem
    Dim insp As Outlook.Inspector
    
    Set insp = Application.ActiveInspector
    
    If insp.CurrentItem.Class = olMail Then
        Set msg = insp.CurrentItem
        
        Dim nextWorkDayDate As Date
        
        Dim nextWorkDay As Integer
        nextWorkDay = Day(DateTime.Now) + 1
        
        Dim dayOfWeek As Integer
        dayOfWeek = Weekday(DateTime.Now, vbSunday)
        
        If dayOfWeek = 5 Then
            nextWorkDay = nextWorkDay + 2 ' Promote from Friday to Sunday
        ElseIf dayOfWeek = 6 Then
            nextWorkDay = nextWorkDay + 1 ' Promote from Satuday to Sunday
        End If
        
        nextWorkDayDate = DateSerial(Year(DateTime.Now), Month(DateTime.Now), nextWorkDay)
        nextWorkDayDate = DateAdd("h", 7, nextWorkDayDate)
        
        msg.DeferredDeliveryTime = nextWorkDayDate
    End If
    
    Set insp = Nothing
    Set msg = Nothing
End Sub

' ------------------------------------------------------------------------------------------------------------------------

