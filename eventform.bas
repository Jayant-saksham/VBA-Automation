Private Sub btnOK_Click()
    ' Validate that all required event details are entered
    If txtStartTime.Text = "" Or txtEndTime.Text = "" Or txtDuration.Text = "" Then
        MsgBox "Please fill in all event details.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Validate that the duration is a positive number
    If Not IsNumeric(txtDuration.Text) Or Val(txtDuration.Text) <= 0 Then
        MsgBox "Please enter a valid positive number for the duration.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Convert duration from minutes to HH:MM format
    Dim durationMinutes As Long
    Dim durationHours As Long
    Dim durationFormatted As String
    
    ' Convert the duration text to a long integer
    durationMinutes = CLng(txtDuration.Text)
    ' Calculate hours and remaining minutes
    durationHours = durationMinutes \ 60
    durationMinutes = durationMinutes Mod 60
    
    ' Format the duration in HH:MM format
    durationFormatted = Format(durationHours, "00") & ":" & Format(durationMinutes, "00")

    ' Define the worksheet to save the event details
    Dim ws As Worksheet
    Set ws = Worksheets("Shift Events")
    
    ' Find the next available row starting from row 5
    Dim nextRow As Long
    nextRow = 5 ' Start checking from row 5

    ' Loop to find the first empty row in column B
    Do While Not IsEmpty(ws.Cells(nextRow, 2))
        nextRow = nextRow + 1
    Loop

    ' Save the event details to the appropriate row in the worksheet
    ws.Cells(nextRow, 2).Value = txtEventName.Text   ' Event Name in Column B
    ws.Cells(nextRow, 4).Value = durationFormatted    ' Duration in HH:MM format in Column D
    ws.Cells(nextRow, 7).Value = txtStartTime.Text    ' Start Time in Column G
    ws.Cells(nextRow, 8).Value = txtEndTime.Text      ' End Time in Column H
    ws.Cells(nextRow, 9).Value = OrganizationalHierarchyForm.txtOrganizationName ' Organization Name in Column I

    ' Display a confirmation message
    MsgBox "Event saved successfully!", vbInformation, "Saved"

    ' Close the EventForm and return to the ScheduleDesignerForm
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    ' Close the EventForm without saving any changes
    Me.Hide
End Sub

' Property to get the Event Duration
Public Property Get eventDuration() As String
    eventDuration = txtDuration.Text
End Property

' Property to get the Event Start Time
Public Property Get eventStartTime() As String
    eventStartTime = txtStartTime.Text
End Property

' Property to get the Event End Time
Public Property Get eventEndTime() As String
    eventEndTime = txtEndTime.Text
End Property
