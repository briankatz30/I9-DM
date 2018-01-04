Imports System.ComponentModel

Public Class Form1

    Dim StrFilePath As String
    Dim Res As New Resizer

    Private Sub RosterAllErrorsViewTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterAllErrorsViewTSMenuItem.Click
        '*************************************************
        ' Popup Menu Item to show All Errors from the Roster Table
        '*************************************************

        ToolStripStatusLabel2.Text = "Displaying All Audit Error Records...."
        Me.Refresh()
        Roster_All_View()

    End Sub

    Private Sub RosterAllQCTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterAllQCTSMenuItem.Click
        '**********************************************
        ' Sub from the popup menu to run the Date QC, SSN QC, 
        '  Other QC when the user selected Roster Table
        '**********************************************
        'Changes the mousepointer to wait until the Grids Load
        Me.Cursor = Cursors.WaitCursor

        Me.Refresh()
        ToolStripStatusLabel2.Text = "Running Date Audit...."
        'Checks Date Errors and Flag them
        Roster_QC_Check_Dates()

        Me.Refresh()
        ToolStripStatusLabel2.Text = "Running SSN Audit...."
        'Check for SSN Errors and Flag them
        Roster_QC_SSN()

        Me.Refresh()
        ToolStripStatusLabel2.Text = "Running Other Field(s) Audit...."
        'Checks for Other Errors and Flags them
        Roster_QC_Required_Fields()

        Me.Refresh()
        ToolStripStatusLabel2.Text = "Completed..."
        'Changes the mousepointer back to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub RosterDateQCTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterDateQCTSMenuItem.Click
        '**************************************************
        ' Sub Routine to open the popup menu to run the Date QC on 
        ' user selected Roster Table
        '**************************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Running Date Audit...."
        Me.Refresh()

        Roster_QC_Check_Dates()

        'Change the cursor back to default
        Me.Cursor = Cursors.Default
        ToolStripStatusLabel2.Text = "Completed...."

    End Sub

    Private Sub RosterDateViewTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterDateViewTSMenuItem.Click
        '********************************************************
        ' Popup Menu Item to show the Date Errors on from the Roster Table
        '********************************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Displaying Date Audit...."
        Me.Refresh()
        Roster_Date_View()

        'Change the cursor back to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub DeleteProjectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteProjectToolStripMenuItem.Click
        '*******************************************************************
        ' Menu Item to Deleting a Project and it all of its tables
        '*******************************************************************
        'Settings before the form Opens
        Form_Project.ProjectComboBox.Visible = True
        Form_Project.ProjectLabel.Text = "Delete a Project"
        Form_Project.AddNewProjectBut.Text = "Delete"
        Form_Project.Text = "Delete a Project"

        'Making the Project Text box not visible so that the user cannot change the name
        Form_Project.ProjectNameTxt.Visible = False

        'Clears out the Form for entry
        Clean_Project_Form()

        'Open for Form
        Form_Project.ShowDialog()

    End Sub

    Private Sub DGVNotMatchTrans_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVNotMatchTrans.CellEnter
        '******************************************************
        ' Sub Routine to View the PDF in the Does Not Match Roster Grid
        '******************************************************
        Dim row As DataGridViewRow = DGVNotMatchTrans.CurrentRow
        Dim I9RosterID As String

        'Need to Grab the the image from I9 if it exsists
        I9RosterID = row.Cells(0).Value.ToString()
        View_Image(I9RosterID, "Transaction")

    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        '************************
        ' Menu Item to Edit a Project
        '************************
        'Form Settings before the Form Opens
        Form_Project.ProjectComboBox.Visible = True
        Form_Project.ProjectLabel.Text = "Project"
        Form_Project.AddNewProjectBut.Text = "Edit"
        Form_Project.Text = "Edit a Project"

        'Making sure the Project Name Text box is visible
        Form_Project.ProjectNameTxt.Visible = True

        'Clears out the Form for entry
        Clean_Project_Form()

        'Opens the Project Form
        Form_Project.ShowDialog()

    End Sub

    Private Sub ExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click
        '*************************************************
        ' Sub Routine to Export the GridView to an Excel Spreadsheet
        '*************************************************
        Me.Cursor = Cursors.WaitCursor
        'Export to Excel from the Menu Item

        Export_Excel(RosterDataGridView)

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub FirstEmployeeBut_Click(sender As Object, e As EventArgs) Handles FirstEmployeeBut.Click
        '**********************************************
        ' Sub Routine to move to the first Record on the Roster
        '**********************************************
        'Get the index of the last row. 
        Dim maxRowIndex As Integer = (RosterDataGridView.Rows.Count - 1 - 1)

        'Compute the index of the current row
        Dim curDataGridViewRow As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim curRowIndex As Integer = curDataGridViewRow.Index

        Dim nextRow As DataGridViewRow = RosterDataGridView.Rows(0)
        'Move the Glyph arrow to the next row
        RosterDataGridView.CurrentCell = nextRow.Cells(0)

        'Select the next row
        nextRow.Selected = True
        RosterDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        RosterDataGridView.CurrentCell.Selected = True

    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        '**********************************
        ' Sub Routine to close any open forms
        ' *********************************
        F_Splash.Close()

    End Sub

    Private Sub Form_Main(sender As Object, e As EventArgs) Handles Me.Load
        '**************************************
        ' Routine to run when the Project Form Loads
        '**************************************
        Res.FindallControls(Me)
        'rs.FindAllControls(Me)



        'Load the Combo Box with the Projects
        Load_Combo_Projects()

        'Status Strip information
        DateTStripStatusLabel.Text = Format(Now(), "d")
        ProjectTStripStatusLabel.Text = "No Project Loaded "
        Me.Refresh()

        'Load a Blank PDF on the form until the user select a row in one grid
        AxAcroPDF1.LoadFile(BlankPDF)

    End Sub

    Private Sub I9DataGridView_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles I9DataGridView.CellEnter
        '**********************************************
        ' Sub Routine to View the PDF in the Transaction Grid
        '**********************************************
        Dim row As DataGridViewRow = I9DataGridView.CurrentRow
        Dim I9RosterID As String

        'Need to Grab the the image from I9 if it exsists
        I9RosterID = row.Cells(0).Value.ToString()
        View_Image(I9RosterID, "Transaction")

    End Sub

    Private Sub ImageI9_Click(sender As Object, e As EventArgs) Handles ImageI9.Click
        '**********************************************
        ' Sub Routine to Display the 1st Page in the Array
        '**********************************************
        CurrentElement = 0

        If CurrentElement >= ImageArray.Count - 1 Then
            Exit Sub
        Else
            'Go to the second page
            AxAcroPDF1.LoadFile(ImageArray(0).ToString)
            TextBox114.Text = ImageArray(CurrentElement).ToString
            ImageSupNext.Enabled = True
            ImageSupPrev.Enabled = True
        End If

    End Sub

    Private Sub ImagePage2_Click(sender As Object, e As EventArgs) Handles ImagePage2.Click
        '**********************************************
        ' Sub Routine to 2nd Page in the Array
        '**********************************************

        'CurrentElement = 1

        'If CurrentElement >= ImageArray.Count - 1 Then
        '    Exit Sub
        'Else
        '    'Go to the second page
        '    AxAcroPDF1.LoadFile(ImageArray(1).ToString)
        '    TextBox114.Text = ImageArray(CurrentElement).ToString
        'End If

    End Sub

    Private Sub ImageSupNext_Click(sender As Object, e As EventArgs) Handles ImageSupNext.Click
        '**********************************************
        ' Sub Routine to move to the next image in the Array
        '**********************************************
        'Move to the next image
        If CurrentElement >= ImageArray.Count - 1 Then
            ImageSupNext.Enabled = False
            ImageSupPrev.Enabled = True
            Exit Sub
        Else
            'Next image in the array
            CurrentElement = CurrentElement + 1
            AxAcroPDF1.LoadFile(ImageArray(CurrentElement).ToString)
            TextBox114.Text = ImageArray(CurrentElement).ToString
            ImageSupPrev.Enabled = True
        End If

    End Sub

    Private Sub ImageSupPrev_Click(sender As Object, e As EventArgs) Handles ImageSupPrev.Click
        '**********************************************
        ' Sub Routine to move to the previous image in the Array
        '**********************************************
        'Move to the previous image
        If CurrentElement = 0 Then
            ImageSupNext.Enabled = True
            ImageSupPrev.Enabled = False
            Exit Sub
        Else
            'Next image in the array
            CurrentElement = CurrentElement - 1
            AxAcroPDF1.LoadFile(ImageArray(CurrentElement).ToString)
            TextBox114.Text = ImageArray(CurrentElement).ToString
            ImageSupPrev.Enabled = True
        End If

    End Sub

    Private Sub LastEmployeeBut_Click(sender As Object, e As EventArgs) Handles LastEmployeeBut.Click
        '**********************************************
        ' Sub Routine to move to the last Record on the Roster
        '**********************************************

        'Get the index of the last row. 
        Dim maxRowIndex As Integer = (RosterDataGridView.Rows.Count - 1 - 1)

        'Compute the index of the current row
        Dim curDataGridViewRow As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim curRowIndex As Integer = curDataGridViewRow.Index

        Dim nextRow As DataGridViewRow = RosterDataGridView.Rows(maxRowIndex)
        'Move the Glyph arrow to the next row
        RosterDataGridView.CurrentCell = nextRow.Cells(0)

        'Select the next row
        nextRow.Selected = True
        RosterDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        RosterDataGridView.CurrentCell.Selected = True

    End Sub

    Private Sub MatchDataGridView_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles MatchDataGridView.CellEnter
        '**********************************************
        ' Sub Routine to show the transcript data in the textboxes
        '**********************************************
        Dim row As DataGridViewRow = MatchDataGridView.CurrentRow

        Dim MatchRow As DataGridViewRow = MatchDataGridView.CurrentRow
        Dim I9RosterID As String

        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        'Need to Grab the the image from I9 if it exsists
        I9RosterID = MatchRow.Cells(0).Value.ToString()
        View_Image(I9RosterID, "Transaction")

        'Displays the data from the Transaction Grid into Text boxes
        TextBox5.Text = row.Cells(12).Value.ToString()              'Employee Status
        TextBox7.Text = row.Cells(13).Value.ToString()              'Alien # LPR
        TextBox8.Text = row.Cells(14).Value.ToString()              'Alien Authorized
        TextBox9.Text = row.Cells(15).Value.ToString()              'Alien Registration
        TextBox28.Text = row.Cells(16).Value.ToString()             'Alien Admission
        TextBox29.Text = row.Cells(17).Value.ToString()             'Foregin Passport
        TextBox27.Text = row.Cells(18).Value.ToString()             'Foregin Country
        TextBox31.Text = row.Cells(27).Value.ToString()             'Employee Signature
        TextBox30.Text = row.Cells(28).Value.ToString()             'Employee Signed Date
        TextBox33.Text = row.Cells(19).Value.ToString()             'Translator Setting
        TextBox32.Text = row.Cells(20).Value.ToString()             'Translator Signature
        TextBox35.Text = row.Cells(21).Value.ToString()             'Translator Last Name
        TextBox34.Text = row.Cells(22).Value.ToString()             'Translator First Name
        TextBox39.Text = row.Cells(29).Value.ToString()             'Translator Sign Date
        TextBox38.Text = row.Cells(23).Value.ToString()             'Translator Address
        TextBox37.Text = row.Cells(24).Value.ToString()             'Translator City
        TextBox36.Text = row.Cells(25).Value.ToString()             'Translator State
        TextBox40.Text = row.Cells(26).Value.ToString()             'Translator Zip

        'Address
        TextBox115.Text = row.Cells(11).Value.ToString()               'Zip Code
        TextBox116.Text = row.Cells(10).Value.ToString()               'State
        TextBox117.Text = row.Cells(9).Value.ToString()                 'City
        TextBox118.Text = row.Cells(8).Value.ToString()                 'Apt #
        TextBox119.Text = row.Cells(7).Value.ToString()                 'Address

        'Section 2
        TextBox45.Text = row.Cells(35).Value.ToString()                 'Issuing Authority
        TextBox46.Text = row.Cells(30).Value.ToString()                 'Last Name
        TextBox47.Text = row.Cells(31).Value.ToString()                 'First Name
        TextBox43.Text = row.Cells(32).Value.ToString()                 'Middle Initial
        TextBox44.Text = row.Cells(33).Value.ToString()                 'Citizenship
        'Document A
        TextBox42.Text = row.Cells(34).Value.ToString()                 'Document Title
        TextBox48.Text = row.Cells(35).Value.ToString()                 'Issuing Authority
        TextBox49.Text = row.Cells(36).Value.ToString()                 'Document Number
        TextBox50.Text = row.Cells(37).Value.ToString()                 'Expiration Date
        'Document B
        TextBox54.Text = row.Cells(38).Value.ToString()                 'Document Title
        TextBox53.Text = row.Cells(39).Value.ToString()                 'Issuing Authority
        TextBox52.Text = row.Cells(40).Value.ToString()                 'Document Number
        TextBox51.Text = row.Cells(41).Value.ToString()                 'Expiration Date
        'Document C
        TextBox58.Text = row.Cells(42).Value.ToString()                 'Document Title
        TextBox57.Text = row.Cells(43).Value.ToString()                 'Issuing Authority
        TextBox56.Text = row.Cells(44).Value.ToString()                 'Document Number
        TextBox55.Text = row.Cells(45).Value.ToString()                 'Expiration Date

        'List  B
        TextBox60.Text = row.Cells(46).Value.ToString()                 'Document Title
        TextBox45.Text = row.Cells(47).Value.ToString()                 'Issuing Authority
        TextBox59.Text = row.Cells(48).Value.ToString()                 'Doucment Number
        TextBox41.Text = row.Cells(49).Value.ToString()                 'Expiration Date

        'List C
        TextBox64.Text = row.Cells(50).Value.ToString()                 'Document Title
        TextBox63.Text = row.Cells(51).Value.ToString()                 'Issuing Authority
        TextBox62.Text = row.Cells(52).Value.ToString()                 'Doucment Number
        TextBox61.Text = row.Cells(53).Value.ToString()                 'Expiration Date

        'Rest of the Text
        TextBox66.Text = row.Cells(54).Value.ToString()             'Additional Information
        TextBox65.Text = row.Cells(55).Value.ToString()             'Employee Start Date
        TextBox68.Text = row.Cells(56).Value.ToString()             'Supervisor Last Name
        TextBox67.Text = row.Cells(57).Value.ToString()             'Supervisor First Name
        TextBox76.Text = row.Cells(58).Value.ToString()             'Supervisor Title
        TextBox71.Text = row.Cells(59).Value.ToString()             'Business Name
        TextBox70.Text = row.Cells(60).Value.ToString()             'Business Address
        TextBox73.Text = row.Cells(61).Value.ToString()             'Business City
        TextBox75.Text = row.Cells(62).Value.ToString()             'Business State
        TextBox74.Text = row.Cells(63).Value.ToString()             'Business Zip
        TextBox69.Text = row.Cells(64).Value.ToString()             'Supervisor Signature
        TextBox72.Text = row.Cells(65).Value.ToString()             'Signed Date

        'Section 3
        TextBox80.Text = row.Cells(66).Value.ToString()             'Employee Last Name
        TextBox77.Text = row.Cells(67).Value.ToString()             'Employee First Name
        TextBox78.Text = row.Cells(68).Value.ToString()             'Employee Middle Initial
        TextBox79.Text = row.Cells(69).Value.ToString()             'Date of Rehire
        TextBox82.Text = row.Cells(70).Value.ToString()             'Document Title
        TextBox81.Text = row.Cells(71).Value.ToString()             'Document Number
        TextBox83.Text = row.Cells(72).Value.ToString()             'Expiration Date
        TextBox85.Text = row.Cells(73).Value.ToString()             'Supervisor Signature
        TextBox84.Text = row.Cells(74).Value.ToString()             'Supervisor Sign Date
        TextBox86.Text = row.Cells(75).Value.ToString()             'Supervisor Name
        TextBox87.Text = row.Cells(76).Value.ToString()             'HandWritten Data
        TextBox89.Text = row.Cells(77).Value.ToString()             'I9 Document Name
        TextBox90.Text = row.Cells(78).Value.ToString()             'I9 Document Name 2
        TextBox88.Text = row.Cells(79).Value.ToString()             'Form Version
        TextBox95.Text = row.Cells(80).Value.ToString()             'I9 Folder

        'Supporting Documents
        TextBox94.Text = row.Cells(81).Value.ToString()             'Supporting Doc 1
        TextBox93.Text = row.Cells(82).Value.ToString()             'Supporting Doc 2
        TextBox92.Text = row.Cells(83).Value.ToString()             'Supoorting Doc 3
        TextBox91.Text = row.Cells(84).Value.ToString()             'Supporting Doc 4
        TextBox96.Text = row.Cells(85).Value.ToString()             'Supporting Doc 5
        TextBox98.Text = row.Cells(86).Value.ToString()             'Supporing Doc 6
        TextBox99.Text = row.Cells(87).Value.ToString()             'Supporing Doc 7
        TextBox100.Text = row.Cells(88).Value.ToString()           'Supporing Doc 8
        TextBox101.Text = row.Cells(89).Value.ToString()           'Supporing Doc 9
        TextBox97.Text = row.Cells(90).Value.ToString()             'Supporing Doc 10
        TextBox103.Text = row.Cells(91).Value.ToString()           'Supporing Doc 11
        TextBox104.Text = row.Cells(92).Value.ToString()           'Supporing Doc 12
        TextBox105.Text = row.Cells(93).Value.ToString()           'Supporing Doc 13
        TextBox106.Text = row.Cells(94).Value.ToString()           'Supporing Doc 14
        TextBox102.Text = row.Cells(95).Value.ToString()           'Supporing Doc 15
        TextBox108.Text = row.Cells(96).Value.ToString()           'Supporing Doc 16
        TextBox109.Text = row.Cells(97).Value.ToString()           'Supporing Doc 17
        TextBox110.Text = row.Cells(98).Value.ToString()           'Supporing Doc 18
        TextBox111.Text = row.Cells(99).Value.ToString()           'Supporing Doc 19
        TextBox107.Text = row.Cells(100).Value.ToString()           'Supporing Doc 20
        TextBox112.Text = row.Cells(101).Value.ToString()          'Supporing Doc 21
        TextBox113.Text = row.Cells(102).Value.ToString()          'Supporing Doc 22

        'Change the cursor to Default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub NextEmployeeBut_Click(sender As Object, e As EventArgs) Handles NextEmployeeBut.Click
        '**********************************************
        ' Sub Routine to move to the next Record on the Roster
        '**********************************************

        'Get the index of the last row. 
        Dim maxRowIndex As Integer = (RosterDataGridView.Rows.Count - 1 - 1)

        'Compute the index of the current row
        Dim curDataGridViewRow As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim curRowIndex As Integer = curDataGridViewRow.Index

        'See if the last row has been passed
        If (curRowIndex >= maxRowIndex) Then
            'The last row has been passed, display an error 
            MsgBox("You are at the last record", vbExclamation, "")
        Else
            Dim nextRow As DataGridViewRow = RosterDataGridView.Rows(curRowIndex + 1)
            'Move the Glyph arrow to the next row
            RosterDataGridView.CurrentCell = nextRow.Cells(0)

            'Select the next row
            nextRow.Selected = True
            RosterDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            RosterDataGridView.CurrentCell.Selected = True

        End If

    End Sub

    Private Sub RosterOtherQCTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterOtherQCTSMenuItem.Click
        '**************************************
        ' Sub Routine to QC other fields on the Roster
        '**************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Running Other Field(s) Audit...."
        Me.Refresh()
        Roster_QC_Required_Fields()

        'Change the cursor to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub OtherShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RosterOtherViewTSMenuItem.Click
        '********************************************
        ' Popup Menu Item to Show Other fields on the Roster
        '********************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Displaying Other Field(s) Audit...."
        Roster_Other_View()

        'Change the cursor to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub PrevEmployeeBut_Click(sender As Object, e As EventArgs) Handles PrevEmployeeBut.Click
        '**********************************************
        ' Sub Routine to move to the Previous Record on the Roster
        '**********************************************

        'Get the index of the last row. 
        Dim maxRowIndex As Integer = (RosterDataGridView.Rows.Count - 1 - 1)

        'Compute the index of the current row
        Dim curDataGridViewRow As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim curRowIndex As Integer = curDataGridViewRow.Index

        'See if the record is the first row
        If (curRowIndex = 0) Or (curRowIndex - 1 < 0) Then
            'The last row has been passed, display an error 
            MsgBox("You are at first record", vbExclamation, "")
        Else
            Dim nextRow As DataGridViewRow = RosterDataGridView.Rows(curRowIndex - 1)
            'Move the Glyph arrow to the next row
            RosterDataGridView.CurrentCell = nextRow.Cells(0)
            'Select the next row
            nextRow.Selected = True
            RosterDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            RosterDataGridView.CurrentCell.Selected = True
        End If

    End Sub

    Private Sub ProjectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewProjectTStripMenu.Click
        '*****************************************************
        ' Menu Item to Open the Project Form for Creating a New Project
        '*****************************************************
        'Settings before opening the Form
        Form_Project.ProjectComboBox.Visible = False
        Form_Project.ProjectLabel.Text = "Create a New Project"
        Form_Project.AddNewProjectBut.Text = "Create New"
        Form_Project.Text = "Create a New Project"

        'Makes sure that the Project Text Box is visible
        Form_Project.ProjectNameTxt.Visible = True

        'Clears out the Form for entry
        Clean_Project_Form()

        'Open the Form
        Form_Project.ShowDialog()

    End Sub

    Private Sub RosterDataGridView_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles RosterDataGridView.CellEndEdit
        '*******************************************************************************
        ' Sub Routine to update Employee ID, Last Name, First Name, Middle and DOB on the roster Grid 
        ' With what the user has selected in the cell in the Database
        '*******************************************************************************
        Dim SqlStmt As String
        Dim row As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = Client_Conn

        Connection = New OleDbConnection(ConnectionString)
        Connection.Open()

        Try
            'Updates the Database with the changes from Grid that the users cell selects
            SqlStmt = "UPDATE [ROSTER] " &
            "Set [Employee Last Name] = '" & row.Cells("Employee Last Name").Value.ToString & "'," &
            "[Employee First Name] = '" & row.Cells("Employee First Name").Value.ToString & "'," &
             "[Employee Middle Name] = '" & row.Cells("Employee Middle Name").Value.ToString & "'," &
            "[Employee SS#] = '" & row.Cells("Employee SS#").Value.ToString & "'," &
            " [Employee Date Of Birth] = '" & row.Cells("Employee Date Of Birth").Value.ToString & "'," &
            " [Employee ID] = '" & row.Cells("Employee ID").Value.ToString & "'" &
            " WHERE ([ROSTER].[ID] = " & row.Cells("ID").Value.ToString & ") ;"

            Dim Cmd2 As OleDbCommand = New OleDbCommand(SqlStmt, Connection)
            Cmd2.ExecuteNonQuery()
            Cmd2.Dispose()

            RosterDataGridView.Refresh()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Connection.Close()
        End Try

    End Sub

    Private Sub RosterDataGridView_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles RosterDataGridView.CellEnter
        '**********************************************
        ' Sub Routine to populate the text boxes from the Roster
        ' Grid view
        '**********************************************
        Dim row As DataGridViewRow = RosterDataGridView.CurrentRow
        Dim ConnectionString, SqlStr As String
        Dim Connection As OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet
        Dim RosterID As String

        'Need this setting some records do have all the fields
        'On Error Resume Next

        'Need to Grab the the image from I9 if it exsists
        RosterID = row.Cells(0).Value.ToString()

        'Load the image if the Transcript Matches the Roster
        View_Image(RosterID, "Roster")

        TextBox1.Text = row.Cells(2).Value.ToString()               'Last Name
        TextBox2.Text = row.Cells(3).Value.ToString()               'First Name
        TextBox3.Text = row.Cells(4).Value.ToString()               'Middle Name
        TextBox4.Text = row.Cells(5).Value.ToString()               'Maiden Name
        TextBox6.Text = row.Cells(1).Value.ToString()               'Employee Number
        TextBox10.Text = row.Cells(9).Value.ToString()             'Street Address
        TextBox11.Text = row.Cells(11).Value.ToString()           'Apt #
        TextBox12.Text = row.Cells(12).Value.ToString()           'City
        TextBox13.Text = row.Cells(13).Value.ToString()           'State
        TextBox14.Text = row.Cells(14).Value.ToString()           'Zip Code
        TextBox15.Text = row.Cells(10).Value.ToString()           'Employee Address # 2
        TextBox16.Text = row.Cells(7).Value.ToString()              'Date of Birth
        TextBox17.Text = row.Cells(8).Value.ToString()              'SSN
        TextBox18.Text = row.Cells(22).Value.ToString()            'Email Address
        TextBox19.Text = row.Cells(18).Value.ToString()            'Home Phone
        TextBox21.Text = row.Cells(20).Value.ToString()            'Cell Phone
        TextBox20.Text = row.Cells(16).Value.ToString()            'Work Phone
        TextBox22.Text = row.Cells(19).Value.ToString()            'Home Ext
        TextBox23.Text = row.Cells(17).Value.ToString()            'Work Ext
        TextBox24.Text = row.Cells(21).Value.ToString()            'Cell Ext
        TextBox25.Text = row.Cells(26).Value.ToString()            'Date of Hire
        TextBox26.Text = row.Cells(27).Value.ToString()            'Term Date

        SqlStr = "Select [ID], [Employee Last Name], [Employee First Name], [Employee Middle Initial], [Employee Maiden Name], [Employee Date Of Birth]," &
             " [Employee SS#], [Employee Address], [Employee Apt #], " &
            " [Employee City], [Employee State], [Employee Zip], [Employee Status], [Employee LPR Alien #], [Employee Alien Authorized To Work Until], " &
            " [Employee Alien Registration/USCIS #], [Employee Alien Or Admission #], [Employee Foreign Passport #], [Employee Foreign Passport Country], " &
            " [Translator Employee Setting], [Translator Signature], [Translator Print Name/Last Name], [Translator First Name], " &
            " [Translator Address], [Translator City], [Translator State], [Translator Zip], [Employee Signature], " &
            " [Employee Signed Date], [Translator Signed Date], [Section 2 Header Last Name], [Section 2 Header First Name], [Section 2 Header Middle Initial]," &
            "[Section 2 Header Immigration Status], [Document Title List A], [Issuing Authority (A)], [Document Number (A1)], [Expiration Date (A1)], " &
            "[Document Title List (A2)], [Issuing Authority (A2)], [Document Number (A2)], [Expiration Date (A2)]," &
            "[Document List (A3)], [Issuing Authority (A3)], [Document Number (A3)], [Expiration Date (A3)], " &
            "[Document Title List (B)], [Issuing Authority (B)], [Document Number (B)], [Expiration Date (B)]," &
            "[Document Title List (C)], [Issuing Authority (C)], [Document Number (C)], [Expiration Date (C)]," &
            "[Section 2 Addition Info], [Employee Start Date], [Supervisor Print/Last Name]," &
            "[Supervisor First Name], [Supervisor Title], [Business Name], [Business Address], [Business City], [Business State]," &
            "[Business Zip], [Supervisor Signature], [Supervisor Signed Date], " &
            "[Employee New Last Name (Section 3)], [Employee First Name (Section 3)], [Employee Middle Initial (Section 3)], [Date of Rehire (Section 3)], " &
            "[Document Title (Section 3)], [Document Number (Section 3)], [Document Expiration Date (Section 3)], [Supervisor Signature (Section 3)], " &
            "[Supervisor Signed Date (Section 3)], [Supervisor Print Name (Section 3)], [Handwritten data in margins] ," &
            "[I-9 Document Name], [I-9 Document Name 2],[Form Version],[I-9 Folder], " &
            "[Supporting Doc 1 Name], [Supporting Doc 2 Name], [Supporting Doc 3 Name], [Supporting Doc 4 Name], [Supporting Doc 5 Name]," &
            "[Supporting Doc 6 Name], [Supporting Doc 7 Name], [Supporting Doc 8 Name], [Supporting Doc 9 Name], [Supporting Doc 10 Name]," &
            "[Supporting Doc 11 Name], [Supporting Doc 12 Name], [Supporting Doc 13 Name], [Supporting Doc 14 Name], [Supporting Doc 15 Name]," &
            "[Supporting Doc 16 Name], [Supporting Doc 17 Name], [Supporting Doc 18 Name], [Supporting Doc 19 Name], [Supporting Doc 20 Name]," &
            "[Supporting Doc 21 Name], [Supporting Doc 22 Name] " &
            "FROM [I9] WHERE [ROSTERID] = '" & RosterID & "' ORDER BY [ID] ;"

        'Connect to the database
        ConnectionString = Client_Conn

        Connection = New OleDbConnection(Client_Conn)
        Connection.Open()
        oledbAdapter = New OleDbDataAdapter(SqlStr, Connection)
        oledbAdapter.Fill(Rs)
        If Rs.Tables(0).Rows.Count > 0 Then
            MatchDataGridView.Visible = True
            'Loads the Grid
            MatchDataGridView.DataSource = Rs.Tables(0)
        Else
            MatchDataGridView.Visible = False
        End If
        Connection.Close()
        Rs.Dispose()

    End Sub

    Private Sub RosterDataGridView_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles RosterDataGridView.CellMouseDown
        '**********************************************
        ' Sub Routine to popup the right click menu for the Roster
        '**********************************************
        If e.Button = MouseButtons.Right Then
            RosterContextMenuStrip.Show(MousePosition)
        End If

    End Sub

    Private Sub RosterImportTSMenu_Click(sender As Object, e As EventArgs) Handles RosterImportTSMenu.Click
        '******************************************************************************
        ' Sub Routine to open the Windows Dialog Box to get the Roster file for import to a Roster table 
        '******************************************************************************

        Try
            'Open the OpenFIleDialog Window
            RosterOpenFileDialog = New System.Windows.Forms.OpenFileDialog
            Dim StrFilePath As String

            'Check to make sure that the user has selected a Project
            If TSProjectComboBox.Text = "" Then
                MsgBox("Please Select a Project before Importing", vbExclamation, "Roster Import")
                Exit Sub
            Else
                'Opens the Windows Dialog box to pick the roster
                With RosterOpenFileDialog
                    .Title = "Select the Tab Delimited Roster File text file to Import"
                    .DefaultExt = ".txt"
                    .Filter = "Roster (Tab Delimited File)|*.txt| All Files |* .*"
                    .FilterIndex = 1
                    If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                        StrFilePath = .FileName
                    End If
                End With
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub RosterOpenFileDialog_FileOk(sender As Object, e As CancelEventArgs) Handles RosterOpenFileDialog.FileOk
        '*****************************************************
        ' Sub Routine that runs when the user selects the file to import.
        '******************************************************
        Dim x As Integer
        Try
            Me.Cursor = Cursors.WaitCursor
            'Load file into the Roster Table
            Load_Roster_File(RosterOpenFileDialog.FileName)
            'Run the Roster Audit
            x = MsgBox("Would you like to run the Audit on the imported File ?", MsgBoxStyle.YesNo, "Run Date Audit")
            If x = 6 Then
                'User wants to run the Audits
                'Runs the Date Audit
                ProjectTStripStatusLabel.Text = "Running Date Audit...."
                Me.Refresh()
                Date_QC_Trans()
                Roster_QC_Check_Dates()
                'Run SSN QC
                ProjectTStripStatusLabel.Text = "Running SSN Audit...."
                Me.Refresh()
                Roster_QC_SSN()
                SSN_QC_Trans()
                ProjectTStripStatusLabel.Text = "Running Other Audit...."
                Me.Refresh()
                Roster_QC_Required_Fields()
                'Run the Match Stored Procedure
                Match_Check()
            Else
                'Do Not run the Audit and Do Nothing
            End If

            MsgBox("Roster Audit Complete", MsgBoxStyle.OkOnly, "")
            TSProjectComboBox_SelectedIndexChanged(Nothing, Nothing)

        Catch ex As Exception
            MsgBox(ex.ToString)

            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub SSNQCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RosterSSNQCTSMenuItem.Click
        '**********************************************
        ' Sub Routine that runs the Roster SSN QC
        '*********************************************
        Me.Cursor = Cursors.WaitCursor
        ProjectTStripStatusLabel.Text = "Running SSN Audit...."

        Me.Refresh()
        Roster_QC_SSN()

        ProjectTStripStatusLabel.Text = "Complete...."
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub SSNShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RosterSSNViewTSMenuItem.Click
        '***************************************************
        ' Popup Menu Item to Show SSN field errors from the Roster
        '***************************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Displaying SSN Audit...."
        Roster_SSN_View()

        'Change the cursor to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub TranscriptImportTSMenu_Click(sender As Object, e As EventArgs) Handles TranscriptImportTSMenu.Click
        '**********************************************
        ' Sub Routine to open the Windows Dialog Box to
        ' Get the Transcription file to import to a project table 
        '**********************************************
        TransOpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Dim StrFilePath As String

        'Check to make sure that the user has selected a Project
        If TSProjectComboBox.Text = "" Then
            MsgBox("Please Select a Project before Importing", vbExclamation, "Transcript Import")
            Exit Sub
        Else
            'Opens the Windows Dialog box to pick the roster
            With TransOpenFileDialog
                .Title = "Select the Transcription File to Import"
                .DefaultExt = ".txt"
                .Filter = "Transcription (Pipe Delimited File)|*.txt| All Files |* .*"
                .FilterIndex = 1
                If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    StrFilePath = .FileName
                End If
            End With
        End If

    End Sub

    Private Sub TransOpenFileDialog_FileOk(sender As Object, e As CancelEventArgs) Handles TransOpenFileDialog.FileOk
        '***************************************************************
        ' Sub Routine that runs when the user selects the transcription file to import.  
        '***************************************************************
        Dim x As Integer

        Try
            Me.Cursor = Cursors.WaitCursor
            'Load text file into Transaction Table
            Load_Trans_File(TransOpenFileDialog.FileName)

            x = MsgBox("Would you like to run an Audit on the imported File ?", MsgBoxStyle.YesNo, "Run Transaction Audit")
            If x = 6 Then
                'User wants to run the Audits

                'Runs the Date Audit
                ProjectTStripStatusLabel.Text = "Running Date Audit"
                Date_QC_Trans()

                'Run SSN QC
                ProjectTStripStatusLabel.Text = "Running SSN Audit"
                SSN_QC_Trans()

                'Run the I9 image check to move the imaged over one field
                ProjectTStripStatusLabel.Text = "Checking for Images in the two I9 Document Name Fields"
                Image_QC()

                'Run the Orphan Doc Check
                ProjectTStripStatusLabel.Text = "Running Orphan Document Audit"
                OrphanRecordCheck()

                'Run the Standalone Section 3 Check
                ProjectTStripStatusLabel.Text = "Running Section 3 Audit"
                StandaloneSection3Check()

                'Run the Translator Signature Check
                ProjectTStripStatusLabel.Text = "Running Translator Signature Audit"
                Translator_QC()

                'Run the Match Stored Procedure
                Match_Check()
            Else
                'Do Nothing
            End If

            MsgBox("Transcription Audit Complete", MsgBoxStyle.OkOnly, "")
            TSProjectComboBox_SelectedIndexChanged(Nothing, Nothing)
            ProjectTStripStatusLabel.Text = "Complete"
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub TSProjectComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TSProjectComboBox.SelectedIndexChanged
        '********************************************************************
        ' Sub Routine to get the Project Name from the Tool Strip Project ComboBox and 
        ' searches the LLX_MAIN_PROJECT Table for the
        ' name of the project to get the ID and other data
        '*********************************************************************

        Dim Connection As OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim ds As New DataTable
        Dim r As DataRow
        Dim SelectedItem As String
        Dim cnn As New ADODB.Connection
        Dim strSql As String

        'Gets what the user has selected from the combobox
        SelectedItem = TSProjectComboBox.SelectedItem.ToString

        'Need to get the selected project table information
        Connection = New OleDbConnection(Main_Conn)
        strSql = "SELECT * FROM MAIN_LLX_PROJECT WHERE [PROJECT_NAME]  = '" & SelectedItem & "' ;"
        Try
            Connection.Open()
            oledbAdapter = New OleDbDataAdapter(strSql, Main_Conn)
            oledbAdapter.Fill(ds)
            GuardianVersion = ""
            For Each r In ds.Rows
                Client_Conn = r("CONNECTION_STRING")
                ProjectID = r("ID")
                ImagePath = r("IMAGE_LOCATION_PATH")
                If r("GUARDIAN_VERSION") IsNot Nothing Then
                    GuardianVersion = r("GUARDIAN_VERSION")
                Else
                    GuardianVersion = ""
                End If
            Next

            oledbAdapter.Dispose()

            'Status Label for the Project Selected
            ProjectTStripStatusLabel.Text = "    Project -  " & SelectedItem & "  |  Guardian Version - " & GuardianVersion
            'Changes the mousepointer to wait until the Grids Load
            Me.Cursor = Cursors.WaitCursor
            'Procedure to Check the Errors of the Roster Table data
            Load_GridViews()
            'Changes the mousepointer back to default
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub I9DataGridView_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles I9DataGridView.CellMouseDown
        '**************************************************
        ' Sub Routine to popup the right click menu for the Transcript
        '**************************************************
        If e.Button = MouseButtons.Right Then
            TranscriptMenuStrip.Show(MousePosition)
        End If

    End Sub

    Private Sub DisplayAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisplayAllToolStripMenuItem.Click
        '*********************************************
        ' Display all the Gridviews from the Menu Item
        '*********************************************
        ToolStripStatusLabel2.Text = "All Records.."
        Me.Refresh()

        Load_GridViews()

    End Sub

    Private Sub DisplayAllToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DisplayAllToolStripMenuItem1.Click
        '*********************************************
        ' Display all the Gridviews from the Menu Item
        '*********************************************
        ToolStripStatusLabel2.Text = "All Records.."
        Me.Refresh()

        Load_GridViews()

    End Sub

    Private Sub DisplayAllToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DisplayAllToolStripMenuItem2.Click
        '*********************************************
        ' Display all the Gridviews from the Menu Item
        '*********************************************
        ToolStripStatusLabel2.Text = "All Records.."
        Me.Refresh()

        Load_GridViews()

    End Sub

    Private Sub OrphanI9sToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrphanI9sToolStripMenuItem.Click
        '***************************************************
        ' Popup Menu Item to Run the Stored Procedure for checking
        ' for Orphan I9s
        '***************************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor
        ToolStripStatusLabel2.Text = "Running Orphan Audit...."
        OrphanRecordCheck()
        Me.Refresh()
        ToolStripStatusLabel2.Text = "Completed...."

        'Change the cursor to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Section3StandaloneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Section3StandaloneToolStripMenuItem.Click
        '***************************************************
        ' Popup Menu Item to Run the Stored Procedure for checking
        ' for Standalone Section 3s
        '***************************************************
        'Change the cursor to wait
        Me.Cursor = Cursors.WaitCursor

        ToolStripStatusLabel2.Text = "Running Section 3 Audit...."
        StandaloneSection3Check()

        ToolStripStatusLabel2.Text = "Completed...."

        'Change the cursor to default
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub DeleteRosterTableMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRosterTableMenuItem.Click
        '***************************************************
        ' Menu Item to Run the Stored Procedure for Dropping the Roster
        ' Table and recreating it with no data
        '***************************************************
        Dim x As Integer
        Dim Selecteditem As String
        'Dim Connection As OleDbConnection
        Dim cmd As New OleDbCommand

        Try
            'Gets what the user has selected
            Selecteditem = TSProjectComboBox.SelectedItem.ToString
            x = MsgBox("Are you sure want to remove ALL Roster data from the Project " & vbCrLf & " " & Selecteditem & "", vbOKCancel + vbExclamation, "Delete Roster Data")
            If x = 1 Then
                'Delete the roster
                Dim Connection = New OleDbConnection(Client_Conn)
                Connection.Open()

                'Need to figure out which Roster table to Drop and recreate
                If GuardianVersion = "G1" Then
                    cmd.CommandText = "dbo.DROP_G1_ROSTER"
                ElseIf GuardianVersion = "G2" Then
                    cmd.CommandText = "dbo.DROP_G1_ROSTER"
                End If

                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = Connection
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                Connection.Dispose()
                x = MsgBox("Roster data has been removed", vbInformation, "Completed")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub ViewSection3TSMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSection3TSMenuItem.Click
        '*****************************************************
        ' Sub routine Menu Item to view Section 3 Standalone Document
        '*****************************************************
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet

        Try
            'Change the cursor to wait
            Me.Cursor = Cursors.WaitCursor

            'Load the Transcription Grid View with the Section 3 Records Only
            ToolStripStatusLabel2.Text = "Loading Transcript Section 3 Data Only...."
            oledbAdapter = New OleDbDataAdapter("SELECT * FROM I9 WHERE [SECTION 3 FLAG] = '1' ORDER BY ID;", Client_Conn)
            oledbAdapter.Fill(Rs)
            I9DataGridView.DataSource = Rs.Tables(0)
            I9DataGridView.Refresh()

            'Housekeeping
            oledbAdapter.Dispose()
            Rs.Dispose()

            'Change the cursor to default
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub ViewOrphanTSMenuItem_Click(sender As Object, e As EventArgs) Handles ViewOrphanTSMenuItem.Click
        '*******************************************
        ' Sub Routine Menu Item to view Orphan Documents
        '*******************************************
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet

        Try
            'Change the cursor to wait
            Me.Cursor = Cursors.WaitCursor

            'Load the Transcription Grid View with the Orphan Doc Records Only
            ToolStripStatusLabel2.Text = "Loading Transcript Section 3 Data Only...."
            oledbAdapter = New OleDbDataAdapter("SELECT * FROM I9 WHERE ORPHANDOC = 'Y' ORDER BY ID;", Client_Conn)
            oledbAdapter.Fill(Rs)
            I9DataGridView.DataSource = Rs.Tables(0)
            I9DataGridView.Refresh()

            'Housekeeping
            oledbAdapter.Dispose()
            Rs.Dispose()

            'Change the cursor to default
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub RosterOverlayTSMenuItem_Click(sender As Object, e As EventArgs) Handles RosterOverlayTSMenuItem.Click
        '******************************************************************
        ' Sub Routine to open the Windows Dialog Box to get an Excel File for the overlay  
        '******************************************************************
        Try
            'Open the OpenFIleDialog Window
            ExcelOpenFileDialog = New System.Windows.Forms.OpenFileDialog
            Dim StrFilePath As String

            'Check to make sure that the user has selected a Project
            If TSProjectComboBox.Text = "" Then
                MsgBox("Please Select a Project before Importing", vbExclamation, "Roster Import")
                Exit Sub
            Else
                'Opens the Windows Dialog box to pick the Excel Spreadsheet
                With ExcelOpenFileDialog
                    .Title = "Select an Excel Spreadsheet to overlay"
                    .DefaultExt = ".xlsx"
                    .Filter = "Excel Files|*.xlsx| All Files |* .*"
                    .FilterIndex = 1
                    If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                        StrFilePath = .FileName
                    End If
                End With
            End If

            'Opens the Mapping Form
            F_Mapping.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    Private Sub ExcelOpenFileDialog_FileOk(sender As Object, e As CancelEventArgs) Handles ExcelOpenFileDialog.FileOk
        '*****************************************************
        ' Sub Routine that runs when the user selects the file to import.
        '*****************************************************
        Dim x As Integer

        Try
            x = MsgBox("You are about to overlay an Excel File in the Roster Table ?", MsgBoxStyle.YesNo, "Run Roster Overlay")
            If x = 6 Then
                'User wants to run the Audits
                ProjectTStripStatusLabel.Text = "Roster Overlay "
                Me.Refresh()
                'Runs the method to process the Excel Spreadsheet
                ImportDataFromExcel(ExcelOpenFileDialog.FileName)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Class
