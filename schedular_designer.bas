Private Sub btnAddEvent_Click()
    ' Show the Event Form to input event details
    EventForm.Show
End Sub

Private Sub btnNext_Click()
    ' Validate the Shift Type input
    If txtShiftType.Text = "" Then
        MsgBox "Please enter the Shift Type.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Validate the Duration input
    If txtDuration.Text = "" Or Not IsNumeric(txtDuration.Text) Then
        MsgBox "Please enter a valid duration (in hours).", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Define the worksheet to work with
    Dim ws As Worksheet
    Set ws = Worksheets("Shift Designer")

    ' Start looking for the next available row from row 12
    Dim nextRow As Long
    nextRow = 12 ' Initialize to row 12

    ' Loop to find the first empty row starting from row 12
    Do While Not IsEmpty(ws.Cells(nextRow, 2)) Or Not IsEmpty(ws.Cells(nextRow, 3))
        nextRow = nextRow + 1
    Loop

    ' Save the inputs to the appropriate row in the worksheet
    ws.Cells(nextRow, 2).Value = txtShiftType.Text   ' Shift Type in Column B
    ws.Cells(nextRow, 3).Value = txtDuration.Text    ' Duration in Column C
    ws.Cells(nextRow, 4).Value = EventForm.txtEventName   ' Event Name from EventForm in Column D
    ws.Cells(nextRow, 5).Value = EventForm.txtDuration    ' Event Duration from EventForm in Column E
    ws.Cells(nextRow, 7).Value = EventForm.txtStartTime   ' Event Start Time from EventForm in Column G
    ws.Cells(nextRow, 8).Value = EventForm.txtEndTime     ' Event End Time from EventForm in Column H
    ws.Cells(nextRow, 9).Value = OrganizationalHierarchyForm.txtOrganizationName ' Organization Name from OrganizationalHierarchyForm in Column I

    ' Display a confirmation message to the user
    MsgBox "Shift saved successfully!", vbInformation, "Saved"

    ' Clear the text boxes for the next entry
    txtShiftType.Text = ""
    txtDuration.Text = ""

    ' Optionally hide the form and move to another step
    Me.Hide
    ShiftEventForm.Show
End Sub

Private Sub btnBack_Click()
    ' Hide this form and show the previous form (OrganizationalHierarchyForm)
    Me.Hide
    OrganizationalHierarchyForm.Show
End Sub

Private Sub btnSave_Click()
    ' Validate the Shift Type input
    If txtShiftType.Text = "" Then
        MsgBox "Please enter the Shift Type.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Validate the Duration input
    If txtDuration.Text = "" Or Not IsNumeric(txtDuration.Text) Then
        MsgBox "Please enter a valid duration (in hours).", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Define the worksheet to work with
    Dim ws As Worksheet
    Set ws = Worksheets("Shift Designer")

    ' Start looking for the next available row from row 12
    Dim nextRow As Long
    nextRow = 12 ' Initialize to row 12

    ' Loop to find the first empty row starting from row 12
    Do While Not IsEmpty(ws.Cells(nextRow, 2)) Or Not IsEmpty(ws.Cells(nextRow, 3))
        nextRow = nextRow + 1
    Loop

    ' Save the inputs to the appropriate row in the worksheet
    ws.Cells(nextRow, 2).Value = txtShiftType.Text   ' Shift Type in Column B
    ws.Cells(nextRow, 3).Value = txtDuration.Text    ' Duration in Column C
    ws.Cells(nextRow, 4).Value = EventForm.txtEventName   ' Event Name from EventForm in Column D
    ws.Cells(nextRow, 5).Value = EventForm.txtDuration    ' Event Duration from EventForm in Column E
    ws.Cells(nextRow, 7).Value = EventForm.txtStartTime   ' Event Start Time from EventForm in Column G
    ws.Cells(nextRow, 8).Value = EventForm.txtEndTime     ' Event End Time from EventForm in Column H
    ws.Cells(nextRow, 9).Value = OrganizationalHierarchyForm.txtOrganizationName ' Organization Name from OrganizationalHierarchyForm in Column I

    ' Display a confirmation message to the user
    MsgBox "Shift saved successfully!", vbInformation, "Saved"

    ' Clear the text boxes for the next entry
    txtShiftType.Text = ""
    txtDuration.Text = ""

    ' Optionally hide the form after saving
    Me.Hide
End Sub
