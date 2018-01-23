
Public Class Form_Project
    Dim da As New OleDbDataAdapter()
    Dim dt As New DataTable()

    Private Sub AddNewProjectBut_Click(sender As Object, e As EventArgs) Handles AddNewProjectBut.Click
        '********************************
        ' Either Adds, Edits or Deletes a Project 
        '********************************
        Dim Rs As New ADODB.Recordset
        Dim RsTable As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection
        Dim ProjectUserSelection, DatabaseNameRosterTable As String
        Dim GVersion, DemoComKey, JIRAText, ProjectDesc, ProductionComKey, ImageLocation As String
        Dim SelectedItem, StrSql, ButtonName As String
        Dim x, c As Integer
        Dim Roster_Connection As String

        'Sets the Connection to the Project Db
        Roster_Connection = Main_Conn

        Try
            If String.IsNullOrEmpty(Trim(ProjectNameTxt.Text).ToString) Then ProjectUserSelection = "" Else ProjectUserSelection = ProjectNameTxt.Text.ToString

            'Check for the Button to be Crete New, Edit or Delete so that it can alter the Project Details
            ButtonName = AddNewProjectBut.Text
            If ButtonName = "Delete" Then ProjectComboBox.Visible = False
            If ButtonName = "Create New" Then ProjectComboBox.Visible = True

            Select Case ButtonName
                Case = "Create New"
                    'Change the Form Values for Add New
                    ProjectLabel.Text = "Create a new Project"
                    AddNewProjectBut.Text = "Create New"

                    'Check the name of the Project before putting Project into the table
                    If Len(ProjectUserSelection) < 1 Or String.IsNullOrEmpty(ProjectUserSelection) Then
                        x = MsgBox("Invalid Project Name", vbOKOnly, "Invalid Entry")
                    Else
                        'Good so far confirm one more time
                        x = MsgBox("Confirm the Project Name ? " & vbCrLf & "                     " & ProjectUserSelection, vbOKCancel, "Create a New Project")
                        If x = 1 Then
                            'Open the Recoredset
                            RosterConnection.Open(Roster_Connection)
                            Rs.Open("SELECT * FROM MAIN_LLX_PROJECT ORDER BY PROJECT_NAME", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

                            If VersionComBox.Text <> String.Empty Then
                                GVersion = VersionComBox.SelectedItem.ToString
                            Else
                                GVersion = ""
                            End If

                            If String.IsNullOrEmpty(DemoComKeyTxt.Text) Then
                                DemoComKey = ""
                            Else
                                DemoComKey = DemoComKeyTxt.Text
                            End If

                            If String.IsNullOrEmpty(ProdComKey.Text) Then
                                ProductionComKey = ""
                            Else
                                ProductionComKey = ProdComKey.Text
                            End If

                            If String.IsNullOrEmpty(JIRATxt.Text) Then
                                JIRAText = ""
                            Else
                                JIRAText = JIRATxt.Text
                            End If

                            If String.IsNullOrEmpty(ProjectDescTx.Text) Then
                                ProjectDesc = ""
                            Else
                                ProjectDesc = ProjectDescTx.Text
                            End If

                            'Add to Main Project Table
                            Rs.AddNew()
                            Rs.Fields.Item(1).Value = ProjectUserSelection.ToString
                            Rs.Fields.Item(2).Value = GVersion
                            Rs.Fields.Item(3).Value = JIRAText
                            Rs.Fields.Item(4).Value = DemoComKey
                            Rs.Fields.Item(5).Value = ProductionComKey
                            Rs.Fields.Item(6).Value = Now()
                            Rs.Fields.Item(7).Value = "Provider=SQLOLEDB.1; Data Source=PHX-035555; Trusted_Connection=True; Initial Catalog= " & ProjectUserSelection.ToString & ";User ID=LLX;Password=Password;"
                            'For JumBox Logon
                            'Rs.Fields.Item(7).Value = "Provider=SQLOLEDB.1;Data Source=PHX-VMJUMP2\SQLEXPRESS;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & ProjectUserSelection.ToString & ";"
                            Rs.Fields.Item(8).Value = ProjectDesc
                            Rs.Fields.Item(9).Value = ImageLocationText.Text
                            Rs.Update()

                            'Needs to run a Stored Procedure to create a database for the project that was just created
                            Dim ConnectionString As String = "Provider=SQLOLEDB.1;Data Source=PHX-035555; Trusted_Connection=True; Initial Catalog= master;User ID=LLX;Password=Password;"
                            Dim NewConnectionString As String = "Provider=SQLOLEDB.1; Data Source=PHX-035555; Trusted_Connection=True; Initial Catalog= " & ProjectUserSelection.ToString & ";User ID=LLX;Password=Password;"

                            'For JumpBox Development
                            'Dim ConnectionString As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=LLX_Project;Data Source=PHX-VMJUMP2\SQLEXPRESS;"
                            'Dim NewConnectionString As String = "Provider=SQLOLEDB.1;Data Source=PHX-VMJUMP2\SQLEXPRESS;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & ProjectUserSelection.ToString & ";"

                            Dim conn As New OleDbConnection(ConnectionString)
                            Dim conn1 As New OleDbConnection(NewConnectionString)
                            Dim ObjReader As System.IO.StreamReader

                            'Need to Create the correct G1 or G2 Roster Table 
                            If GVersion = "G1" Then
                                'Create the Database with the G1 table format
                                ObjReader = New System.IO.StreamReader(Application.StartupPath & "\DbScripts\CREATEDBTABLESG1.sql")
                            Else    'Will create the G2 with anyother choice
                                'Creates the New Database and tables
                                ObjReader = New System.IO.StreamReader(Application.StartupPath & "\DbScripts\CREATEDBTABLESG2.sql")
                            End If

                            Dim CreateScript As String = ObjReader.ReadToEnd
                            'Replace the DbProjectName with the Project Name in the text file
                            CreateScript = Replace(CreateScript, "DbProjectName", ProjectUserSelection.ToString)
                            ObjReader.Close()
                            Dim cmd = New OleDbCommand(CreateScript, conn)
                            conn.Open()
                            cmd.ExecuteNonQuery()
                            conn.Close()

                            'Creates the Stored Procedures for the newly created Db
                            ObjReader = New System.IO.StreamReader(Application.StartupPath & "\DbScripts\CREATE_STORED_PROCEDURES.sql")
                            Dim NewCase As String = ObjReader.ReadToEnd
                            conn1.Open()
                            cmd = New OleDbCommand(NewCase, conn1)
                            cmd.ExecuteNonQuery()
                            conn1.Close()

                            'Housekeeping
                            conn = Nothing
                            conn1 = Nothing
                            ObjReader = Nothing
                            Rs.Close()

                            'Displays the Message Box that the project was added to the Project Table
                            c = MsgBox("Project was created successfully....", vbOKOnly, "Created Successfully")
                        End If
                    End If

                Case = "Edit"
                    'Checks for Nulls before saving the data back to the MAIN LLX PROJECT table
                    ProjectNameTxt.Visible = False

                    If VersionComBox.Text <> String.Empty Then
                        GVersion = VersionComBox.SelectedItem.ToString
                    Else
                        GVersion = ""
                    End If

                    If String.IsNullOrEmpty(DemoComKeyTxt.Text) Then
                        DemoComKey = ""
                    Else
                        DemoComKey = DemoComKeyTxt.Text
                    End If

                    If String.IsNullOrEmpty(ProdComKey.Text) Then
                        ProductionComKey = ""
                    Else
                        ProductionComKey = ProdComKey.Text
                    End If

                    If String.IsNullOrEmpty(JIRATxt.Text) Then
                        JIRAText = ""
                    Else
                        JIRAText = JIRATxt.Text
                    End If

                    If String.IsNullOrEmpty(ProjectDescTx.Text) Then
                        ProjectDesc = ""
                    Else
                        ProjectDesc = ProjectDescTx.Text
                    End If

                    ImageLocation = ImageLocationText.Text

                    If ProjectDescTx.Text <> String.Empty Then
                        ProjectDesc = ProjectDescTx.Text
                    Else
                        ProjectDesc = ""
                    End If

                    'Gets what the user has selected
                    SelectedItem = ProjectComboBox.SelectedItem.ToString

                    'Need to get all the project table information
                    StrSql = "Select * FROM [MAIN_LLX_PROJECT] WHERE [PROJECT_NAME]  = '" & SelectedItem & "' ORDER BY PROJECT_NAME ;"
                    RosterConnection.Open(Roster_Connection)
                    Rs.Open(StrSql, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

                    'Update the Values
                    Rs.Fields.Item(1).Value = ProjectUserSelection.ToString
                    Rs.Fields.Item(2).Value = GVersion
                    Rs.Fields.Item(3).Value = JIRAText
                    Rs.Fields.Item(4).Value = DemoComKey
                    Rs.Fields.Item(5).Value = ProductionComKey
                    Rs.Fields.Item(6).Value = Now()
                    Rs.Fields.Item(7).Value = "Provider = SQLOLEDB.1; Data Source=PHX-035555; Trusted_Connection=True; Initial Catalog=" & ProjectUserSelection.ToString & ";User ID=LLX;Password=Password;"
                    'For JumpBox Development
                    'Rs.Fields.Item(7).Value = "Provider=SQLOLEDB.1;Data Source=PHX-VMJUMP2\SQLEXPRESS;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & ProjectUserSelection.ToString & ";"
                    Rs.Fields.Item(8).Value = ProjectDesc
                    Rs.Fields.Item(9).Value = ImageLocationText.Text
                    Rs.Update()
                    Rs.Close()

                    'Displays the Message Box that the project was saved to the Project Table
                    c = MsgBox("Project was saved successfully....", vbOKOnly, "Saved Successfully")

                Case = "Delete"
                    'Gets what the user has selected
                    SelectedItem = ProjectComboBox.SelectedItem.ToString
                    x = MsgBox("Are you sure want to delete this Project and all of the data " & vbCrLf & " associated with it (including the Roster and Transcription data) ? ", vbOKCancel + vbExclamation, "Delete Project")
                    If x = 1 Then
                        'Delete this sucker
                        StrSql = "SELECT * FROM MAIN_LLX_PROJECT WHERE [PROJECT_NAME]  = '" & SelectedItem & "' ORDER BY [PROJECT_NAME] ;"
                        'Need to get all the Drop the Db Project along with Roster and Transcript
                        DatabaseNameRosterTable = SelectedItem
                        Dim Connection As OleDbConnection = New OleDbConnection(Main_Conn)
                        Connection.Open()
                        Dim SQLStmt As String = "DROP DATABASE [" & DatabaseNameRosterTable & "]"
                        Dim Cmd1 As OleDbCommand = New OleDbCommand(SQLStmt, Connection)
                        Cmd1.ExecuteNonQuery()

                        'Delete the Project Row from the Main Project table
                        SQLStmt = "DELETE FROM [MAIN_LLX_PROJECT] WHERE [PROJECT_NAME] = '" & SelectedItem & "'"
                        Cmd1 = New OleDbCommand(SQLStmt, Connection)
                        Cmd1.ExecuteNonQuery()
                        Cmd1.Dispose()
                        Connection.Dispose()
                    Else
                        'User Selected to Cancel
                    End If
            End Select

            'Need to reload the Combo Box
            Load_Combo_Projects()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

        'Closes the Project Form
        Me.Close()

    End Sub

    Private Sub CancelProjectCmd_Click(sender As Object, e As EventArgs) Handles CancelProjectCmd.Click
        '**********************************************
        ' Cancel Command Button
        '**********************************************
        'Closes the Project Form
        Me.Close()

    End Sub

    Private Sub Form_Project_Load(sender As Object, e As EventArgs) Handles Me.Load
        '*********************************************************************
        ' Loads the public Variable for the path of the images into the TextBox for the user
        ' ********************************************************************
        ImageLocationText.Text = ServerShare

    End Sub

    Private Sub ProjectComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ProjectComboBox.SelectedIndexChanged
        '***************************************************************************************
        ' Sub Routine to get the Project Name from the Project  ComboBox and then Load the fields with the data to 
        ' T_Project Table and Edit them or Cancel
        '****************************************************************************************
        Dim SelectedItem As String
        Dim strSql As String
        Dim Rs As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection

        RosterConnection.Open(Main_Conn)

        'Gets what the user has selected
        SelectedItem = ProjectComboBox.SelectedItem.ToString

        'Need to get all the project table information
        strSql = "Select * FROM [MAIN_LLX_PROJECT] WHERE [PROJECT_NAME]  = '" & SelectedItem & "' ORDER BY [PROJECT_NAME] ;"
        Try
            Rs.Open(strSql, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            'Displays the Project in the Text Boxes
            Do While Not Rs.EOF
                ProjectNameTxt.Text = Rs.Fields.Item("PROJECT_NAME").Value
                VersionComBox.Text = Rs.Fields.Item("GUARDIAN_VERSION").Value
                DemoComKeyTxt.Text = Rs.Fields.Item("DEMO_COM_KEY").Value
                ProdComKey.Text = Rs.Fields.Item("PROD_COM_KEY").Value
                JIRATxt.Text = Rs.Fields.Item("JIRA_NUMBER").Value
                ImageLocationText.Text = Rs.Fields.Item("IMAGE_LOCATION_PATH").Value
                ProjectDescTx.Text = Rs.Fields.Item("DESCRIPTION").Value
                Rs.MoveNext()
            Loop

            'HouseKeeping
            Rs.Close()
            RosterConnection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Class