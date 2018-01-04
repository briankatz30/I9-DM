Module M_Other
    Public Sub Roster_QC_Required_Fields()
        '**************************************
        ' Sub to QC Other Fields from the Roster Table
        '**************************************
        Dim ErrorMsg, OtherDesc As String
        Dim RsRoster As New ADODB.Recordset
        Dim RsWipe As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String

        Roster_Connection = Client_Conn

        Try
            'Need to wipe the Other columns before we start
            RosterConnection.Open(Roster_Connection)
            RsWipe.Open("SELECT [OTHER ERROR], [OTHER DESCRIPTION] FROM [ROSTER] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RsWipe.MoveFirst()
            Do While Not RsWipe.EOF
                RsWipe.Fields.Item("OTHER ERROR").Value = DBNull.Value.ToString
                RsWipe.Fields.Item("OTHER DESCRIPTION").Value = DBNull.Value.ToString
                RsWipe.Update()
                RsWipe.MoveNext()
            Loop
            RsWipe.Close()
            RosterConnection.Close()

            ErrorMsg = ""
            RosterConnection.Open(Roster_Connection)
            'Employee ID null check
            RsRoster.Open("SELECT [ID], [OTHER ERROR], [OTHER DESCRIPTION] FROM [ROSTER] WHERE ([EMPLOYEE ID] IS NULL) OR [EMPLOYEE ID] = '';", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RsRoster.EOF Then
                'All records have an ID update the Import Box
            Else
                'Loop through the Empoyees with No IDs
                Do While Not RsRoster.EOF
                    OtherDesc = If(IsDBNull(RsRoster.Fields.Item("OTHER DESCRIPTION").Value), String.Empty, RsRoster.Fields.Item("OTHER DESCRIPTION").Value)
                    If OtherDesc = "" Then
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = OtherDesc & "Missing Employee ID : "
                        RsRoster.Update()
                    Else
                        ErrorMsg = RsRoster.Fields.Item("OTHER DESCRIPTION").Value
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = ErrorMsg & "Missing Employee ID : "
                        RsRoster.Update()
                    End If
                    RsRoster.MoveNext()
                Loop
            End If

            RsRoster.Close()
            RosterConnection.Close()
            RosterConnection.Open(Roster_Connection)

            'Need to check for Employee First Name
            RsRoster.Open("SELECT [ID], [OTHER ERROR], [OTHER DESCRIPTION] FROM [ROSTER] WHERE ([EMPLOYEE FIRST NAME] IS NULL) " &
            "Or [EMPLOYEE ID] = '';", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RsRoster.EOF Then
                'All records have a First Name
            Else
                'Loop through the Empoyees with No First Name
                Do While Not RsRoster.EOF
                    OtherDesc = If(IsDBNull(RsRoster.Fields.Item("OTHER DESCRIPTION").Value), String.Empty, RsRoster.Fields.Item("OTHER DESCRIPTION").Value)
                    If OtherDesc = "" Then
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = OtherDesc & "Missing Employee First Name : "
                        RsRoster.Update()
                    Else
                        ErrorMsg = RsRoster.Fields.Item("OTHER DESCRIPTION").Value
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = ErrorMsg & "Missing Employee Employee First Name : "
                        RsRoster.Update()
                    End If
                    RsRoster.MoveNext()
                Loop
            End If

            RsRoster.Close()
            RosterConnection.Close()
            RosterConnection.Open(Roster_Connection)
            'Need to check for Employee Last Name
            RsRoster.Open("SELECT [ID], [OTHER ERROR], [OTHER DESCRIPTION] FROM [ROSTER] WHERE ([EMPLOYEE LAST NAME] IS NULL) " &
            " OR [EMPLOYEE LAST NAME] = '';", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RsRoster.EOF Then
                'All records have a Last Name
            Else
                'Loop through the Empoyees with No Last Name
                Do While Not RsRoster.EOF
                    OtherDesc = If(IsDBNull(RsRoster.Fields.Item("OTHER DESCRIPTION").Value), String.Empty, RsRoster.Fields.Item("OTHER DESCRIPTION").Value)
                    If OtherDesc = "" Then
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = OtherDesc & "Missing Employee Last Name : "
                        RsRoster.Update()
                    Else
                        ErrorMsg = RsRoster.Fields.Item("OTHER DESCRIPTION").Value
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = ErrorMsg & "Missing Employee Employee Last Name : "
                        RsRoster.Update()
                    End If
                    RsRoster.MoveNext()
                Loop
            End If

            RsRoster.Close()
            RosterConnection.Close()
            RosterConnection.Open(Roster_Connection)
            'Need to check for Location Name \ Business Unit are blank
            RsRoster.Open("SELECT [ID] , [OTHER ERROR], [OTHER DESCRIPTION] FROM  " &
            "[ROSTER]  WHERE ([Location Name] Is Null OR [LOCATION NAME] = '') " &
            "AND ([Location Number] Is Null OR [LOCATION NUMBER] = '') AND ([BUSINESS UNIT] Is NULL OR [BUSINESS UNIT] = '');", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RsRoster.EOF Then
                'All records have a Location Name \ Location
            Else
                'Loop through the Empoyees with No Location Name \ Location
                Do While Not RsRoster.EOF
                    OtherDesc = If(IsDBNull(RsRoster.Fields.Item("OTHER DESCRIPTION").Value), String.Empty, RsRoster.Fields.Item("OTHER DESCRIPTION").Value)
                    If OtherDesc = "" Then
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = OtherDesc & " Missing Location Name \ Number or Business Unit : "
                        RsRoster.Update()
                    Else
                        ErrorMsg = RsRoster.Fields.Item("OTHER DESCRIPTION").Value
                        RsRoster.Fields.Item("Other Error").Value = True
                        RsRoster.Fields.Item("Other Description").Value = ErrorMsg & " Missing Location Name \ Number or Business Unit : "
                        RsRoster.Update()
                    End If
                    RsRoster.MoveNext()
                Loop
            End If
            RsRoster.Close()

            If GuardianVersion = "G1" Then
                RosterConnection.Close()
                RosterConnection.Open(Roster_Connection)
                'Need to check for Occuption Class
                RsRoster.Open("SELECT [ID], [OTHER ERROR], [OTHER DESCRIPTION] FROM [ROSTER]  WHERE ([OCCUPATION CLASS] Is NULL OR [OCCUPATION CLASS = '') " &
                ";", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                If RsRoster.EOF Then
                    'All records have a Occuption Class
                Else
                    'Loop through the Empoyees with No Occuption Class
                    Do While Not RsRoster.EOF
                        OtherDesc = If(IsDBNull(RsRoster.Fields.Item("OTHER DESCRIPTION").Value), String.Empty, RsRoster.Fields.Item("OTHER DESCRIPTION").Value)
                        If OtherDesc = "" Then
                            RsRoster.Fields.Item("Other Error").Value = True
                            RsRoster.Fields.Item("Other Description").Value = OtherDesc & " Missing Occuption Class :  "
                            RsRoster.Update()
                        Else
                            ErrorMsg = RsRoster.Fields.Item("OTHER DESCRIPTION").Value
                            RsRoster.Fields.Item("Other Error").Value = True
                            RsRoster.Fields.Item("Other Description").Value = ErrorMsg & "Missing Occuption Class : "
                            RsRoster.Update()
                        End If
                        RsRoster.MoveNext()
                    Loop
                    RsRoster.Close()
                End If
            End If

            'Housekeeping
            RosterConnection.Close()
            RosterConnection = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Sub Roster_Other_View()
        '*******************************************************************************
        ' Menu Item to Change the the Roster Grid View to show just Other Errors From the Roster Table
        '*******************************************************************************

        Dim ConnectionString As String
        Dim x As Integer
        Dim SqlStr As String
        Dim Connection As OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet

        'Connect to the database
        ConnectionString = Client_Conn

        'Need to Connect to the Db
        Connection = New OleDbConnection(ConnectionString)
        Connection.Open()

        'SqlStr = "Select [ID], [EMPLOYEE ID], [EMPLOYEE LAST NAME], [EMPLOYEE FIRST NAME], [EMPLOYEE MIDDLE NAME], " &
        '" [EMPLOYEE DATE OF BIRTH], [EMPLOYEE SS#], [LOCATION NAME], [LOCATION NUMBER], [HIRE DATE], [TERMINATED DATE], " &
        '" [DATE ERROR], [DATE DESCRIPTION], [SSN ERROR], [SSN DESCRIPTION], [OTHER ERROR], [OTHER DESCRIPTION] " &
        '" FROM [ROSTER] WHERE [OTHER ERROR] = '-1' ORDER BY [ID] ;"

        SqlStr = "SELECT * FROM [ROSTER] WHERE [OTHER ERROR] = '-1' ORDER BY [ID] ;"
        oledbAdapter = New OleDbDataAdapter(SqlStr, ConnectionString)
        oledbAdapter.Fill(Rs)

        'Loads the Grid with the SQL results
        Form1.RosterDataGridView.DataSource = Rs.Tables(0)

        'Displays the Record Count
        x = Form1.RosterDataGridView.Rows.Count
        Form1.StatusLabelGridCount.Text = " Roster Records - " & Format(x, "#,###")

        'HouseKeeping
        Connection.Close()
        oledbAdapter.Dispose()

    End Sub

    Public Sub Roster_All_View()
        '************************************************************************
        ' Menu Item to Change the the Roster Grid View to show All Errors From the Roster Table
        '************************************************************************

        Dim ConnectionString As String
        Dim SqlStr As String
        Dim Connection As OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet
        Dim x As Integer

        'Connect to the database
        ConnectionString = Client_Conn

        'Need to Connect to the Db
        Connection = New OleDbConnection(ConnectionString)
        Connection.Open()

        SqlStr = "Select [ID], [EMPLOYEE ID], [EMPLOYEE LAST NAME], [EMPLOYEE FIRST NAME], [EMPLOYEE MIDDLE NAME], " &
        " [EMPLOYEE Date Of BIRTH], [EMPLOYEE SS#], [LOCATION NAME], [LOCATION NUMBER], [HIRE Date], [TERMINATED Date], " &
        " [DATE Error], [DATE DESCRIPTION], [SSN Error], [SSN DESCRIPTION], [OTHER Error], [OTHER DESCRIPTION] " &
        " FROM [ROSTER] WHERE [DATE Error] = '-1' OR [SSN ERROR] = '-1' OR [OTHER ERROR] = '-1' ORDER BY [DATE ERROR] DESC,[SSN ERROR] DESC, [OTHER ERROR] DESC ;"

        oledbAdapter = New OleDbDataAdapter(SqlStr, ConnectionString)
        oledbAdapter.Fill(Rs)

        'Loads the Grid
        Form1.RosterDataGridView.DataSource = Rs.Tables(0)

        'Displays the Record Count
        x = Form1.RosterDataGridView.Rows.Count - 1
        Form1.StatusLabelGridCount.Text = " Roster Records - " & Format(x, "#,###")

        'HouseKeeping
        Connection.Close()
        oledbAdapter.Dispose()

    End Sub

    Public Sub OrphanRecordCheck()
        '**********************************************
        ' Menu Item for checking for Ophan Records and updating
        ' the I9 table for review
        '**********************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "dbo.ORPHANDOCS"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Sub StandaloneSection3Check()
        '**********************************************
        ' Menu Item for checking for Checking for Standalone Section3s
        ' and updating the I9 table for review
        '**********************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "dbo.FINDSECTION3S"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Translator_QC()
        '**********************************************
        ' Menu Item for checking for Checking the Translator Signature
        ' and updating the I9 table for review
        '**********************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "dbo.TRANSLATOR_QC"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Image_QC()
        '*****************************************************
        ' Menu Item for checking Images on the I9 Document Name
        ' and I9 Document Name 2 and updating the I9 table for review
        '*****************************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "dbo.UPDATEIMAGELOCATIONS"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub ViewStandaloneSection3()
        '**************************************************
        ' Menu Item for checking for Viewing Standalone Section3s
        '**************************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "SELECT * FROM I9 WHERE [SECTION 3 FLAG] = 1"
            cmd.CommandType = CommandType.Text
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()



        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Module
