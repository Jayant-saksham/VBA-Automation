Private Sub UserForm_Initialize()
    ' Populate the combo box with hierarchy levels when the form initializes
    cboHierarchyLevel.AddItem "1"
    cboHierarchyLevel.AddItem "2"
    cboHierarchyLevel.AddItem "3"
    cboHierarchyLevel.AddItem "4"
    cboHierarchyLevel.AddItem "5"
End Sub

Private Sub btnNext_Click()
    ' Validate the Organization Name input
    If txtOrganizationName.Text = "" Then
        MsgBox "Please enter the Organization Name.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Validate that a Hierarchy Level has been selected
    If cboHierarchyLevel.Value = "" Then
        MsgBox "Please select a Hierarchy Level.", vbExclamation, "Input Error"
        Exit Sub
    End If

    ' Define the worksheet to work with
    Dim ws As Worksheet
    Set ws = Worksheets("Organizational Hierarchy")

    ' Determine the next available row starting from row 5
    Dim nextRow As Long
    nextRow = 5

    ' Loop to find the first empty row in columns B, C, D, and E
    Do While Not IsEmpty(ws.Cells(nextRow, 2)) Or Not IsEmpty(ws.Cells(nextRow, 3)) Or Not IsEmpty(ws.Cells(nextRow, 4)) Or Not IsEmpty(ws.Cells(nextRow, 5))
        nextRow = nextRow + 1
    Loop

    ' Generate the Tree Structure View based on the selected hierarchy level
    Dim treeStructure As String
    treeStructure = String(cboHierarchyLevel.Value, "-") & txtOrganizationName.Text

    ' Save the input values to the appropriate columns in the worksheet
    ws.Cells(nextRow, 2).Value = cboHierarchyLevel.Value   ' Hierarchy Level in Column B
    ws.Cells(nextRow, 3).Value = txtOrganizationName.Text  ' Organization Name in Column C
    ws.Cells(nextRow, 4).Value = treeStructure             ' Tree Structure in Column D
    ws.Cells(nextRow, 5).Value = "Enter definition here"   ' Placeholder text in Column E

    ' Display a confirmation message to the user
    MsgBox "Hierarchy Level saved successfully!", vbInformation, "Saved"

    ' Hide this form and show the next form
    Me.Hide
    ScheduleDesignerForm.Show
End Sub

Private Sub btnBack_Click()
    ' Unload the current form when the Back button is clicked
    Unload Me
End Sub

Private Sub btnSave_Click()
    ' Define the worksheet to work with
    Dim ws As Worksheet
    Set ws = Worksheets("Organizational Hierarchy")
    
    ' Determine the next available row starting from row 5
    Dim nextRow As Long
    nextRow = 5
    Do While ws.Cells(nextRow, 2).Value <> ""
        nextRow = nextRow + 1
    Loop
    
    ' Generate the Tree Structure View based on the selected hierarchy level
    Dim treeStructure As String
    treeStructure = String(cboHierarchyLevel.Value, "-") & txtOrganizationName.Text

    ' Save the input values to the next available row in the worksheet
    ws.Cells(nextRow, 2).Value = cboHierarchyLevel.Value  ' Hierarchy Level in Column B
    ws.Cells(nextRow, 3).Value = txtOrganizationName.Text ' Organization Name in Column C
    ws.Cells(nextRow, 4).Value = treeStructure            ' Tree Structure in Column D
    ws.Cells(nextRow, 5).Value = "Enter definition here"  ' Placeholder text in Column E

    ' Notify the user that the inputs have been saved successfully
    MsgBox "Inputs saved successfully", vbInformation, "Saved"
End Sub
