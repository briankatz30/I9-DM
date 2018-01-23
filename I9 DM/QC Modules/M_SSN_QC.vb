Module M_SSN_QC

    Public Sub Roster_QC_SSN()
        '**************************************
        ' Sub to QC SSN numbers in the Roster Table
        '**************************************
        Dim RosterConnection As New ADODB.Connection
        Dim RosterSSNConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Dim RosterSSN As New ADODB.Recordset
        Dim RosterFix As New ADODB.Recordset
        Dim RsDupCheck As New ADODB.Recordset
        Dim Rs As New ADODB.Recordset

        Dim SSN, FormatSSN, DUPSSN As String
        Dim ErrorMsg, SSNDesc As String
        Dim Post As Integer

        Roster_Connection = Client_Conn

        Form1.Cursor = Cursors.WaitCursor

        Try
            'Need to wipe the SSN columns before we start
            RosterConnection.Open(Roster_Connection)
            RosterFix.Open("SELECT [SSN Error], [SSN Description] FROM [ROSTER] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RosterFix.MoveFirst()
            Do While Not RosterFix.EOF
                RosterFix.Fields.Item("SSN ERROR").Value = DBNull.Value.ToString
                RosterFix.Fields.Item("SSN DESCRIPTION").Value = DBNull.Value.ToString
                RosterFix.MoveNext()
            Loop
            RosterFix.Close()

            ErrorMsg = ""
            FormatSSN = ""

            RosterSSN.Open("SELECT [ID], [EMPLOYEE SS#],  [SSN ERROR], [SSN DESCRIPTION] " &
            " FROM [ROSTER] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RosterSSN.MoveFirst()

            'Loop through the Recordset
            Do While Not RosterSSN.EOF
                ErrorMsg = ""
                SSN = If(IsDBNull(RosterSSN.Fields.Item("EMPLOYEE SS#").Value), String.Empty, RosterSSN.Fields.Item("EMPLOYEE SS#").Value)
                SSNDesc = If(IsDBNull(RosterSSN.Fields.Item("SSN DESCRIPTION").Value), String.Empty, RosterSSN.Fields.Item("SSN DESCRIPTION").Value)
                'No SSN Log the Error
                If SSN = "" Then
                    If SSNDesc = "" Then
                        ErrorMsg = ""
                    Else
                        ErrorMsg = RosterSSN.Fields.Item("SSN DESCRIPTION").Value
                    End If
                    'Write to the Description field
                    RosterSSN.Fields.Item("SSN ERROR").Value = True
                    RosterSSN.Fields.Item("SSN DESCRIPTION").Value = ErrorMsg & " SSN is blank : "
                    RosterSSN.Update()
                Else
                    'Check for the correct format of the field
                    If Len(SSN) = 11 And Post = InStr(4, SSN, "-") = 0 Then
                        'Has the right amount of characters and dashes
                        FormatSSN = SSN
                    ElseIf Len(SSN) = 9 Then
                        'If the length is 9 and not less than need to place the - in the right place
                        FormatSSN = Left(SSN, 3) & "-" & Mid(SSN, 4, 2) & "-" & Right(SSN, 4)
                    Else
                        If SSNDesc = "" Then
                            ErrorMsg = ""
                        Else
                            ErrorMsg = RosterSSN.Fields.Item("SSN DESCRIPTION").Value
                        End If
                        'Write to the Description field
                        RosterSSN.Fields.Item("SSN ERROR").Value = True
                        RosterSSN.Fields.Item("SSN DESCRIPTION").Value = ErrorMsg & " SSN is Not a valid number of characters  : "
                        RosterSSN.Update()
                    End If
                    'Write the valid SSN Number Format the field
                    RosterSSN.Fields.Item("Employee SS#").Value = FormatSSN
                    RosterSSN.Update()
                End If

                'Check for invalid numbers in the SSN
                If Left(FormatSSN, 4) = "000" Or Mid(FormatSSN, 5, 2) = "00" Or Right(FormatSSN, 4) = "0000" _
                Or Left(FormatSSN, 3) = "666" Or Left(FormatSSN, 1) = "9" Then
                    'Invalid SSN Number log the error
                    SSNDesc = If(IsDBNull(RosterSSN.Fields.Item("SSN DESCRIPTION").Value), String.Empty, RosterSSN.Fields.Item("SSN DESCRIPTION").Value)
                    If SSNDesc = "" Then
                        ErrorMsg = ""
                    Else
                        ErrorMsg = RosterSSN.Fields.Item("SSN DESCRIPTION").Value
                    End If
                    'Write to the Description field
                    RosterSSN.Fields.Item("SSN ERROR").Value = True
                    RosterSSN.Fields.Item("SSN DESCRIPTION").Value = ErrorMsg & " SSN is Not a valid number : "
                    RosterSSN.Update()
                Else
                    'Write the valid SSN Number Format the field
                    RosterSSN.Fields.Item("Employee SS#").Value = FormatSSN
                    RosterSSN.Update()
                End If
                RosterSSN.MoveNext()
            Loop
            RosterSSN.Close()

            RosterSSNConnection.Open(Roster_Connection)

            'Check for Duplicate SSN Numbers
            RsDupCheck.Open("SELECT [Employee SS#], COUNT([Employee SS#]) FROM [ROSTER] GROUP BY " &
            "[Employee SS#] HAVING (COUNT([Employee SS#]) > 1);", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RsDupCheck.EOF Then
                'Means no dups in the Roster
            Else
                RsDupCheck.MoveFirst()

                'Loop Through the Dup SSN numbers
                Do While Not RsDupCheck.EOF
                    ErrorMsg = ""
                    DUPSSN = If(IsDBNull(RsDupCheck.Fields.Item("EMPLOYEE SS#").Value), String.Empty, RsDupCheck.Fields.Item("EMPLOYEE SS#").Value)
                    If DUPSSN = "" Then
                        'Blank SSN Exit to the main loop
                    Else
                        DUPSSN = RsDupCheck.Fields.Item("EMPLOYEE SS#").Value
                        Rs.Open("SELECT [SSN DESCRIPTION], [SSN ERROR] FROM [ROSTER] WHERE " &
                         " [EMPLOYEE SS#] = '" & DUPSSN & "' ;", RosterSSNConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                        Do While Not Rs.EOF
                            'Reset the Error Message
                            ErrorMsg = ""
                            'Log the Dup error
                            SSNDesc = If(IsDBNull(Rs.Fields.Item("SSN DESCRIPTION").Value), String.Empty, Rs.Fields.Item("SSN DESCRIPTION").Value)
                            If SSNDesc = "" Then
                                Rs.Fields.Item("SSN Error").Value = True
                                Rs.Fields.Item("SSN Description").Value = ErrorMsg & "Duplicate SSN - " & DUPSSN & " : "
                                Rs.Update()
                            Else
                                Rs.Fields.Item("SSN Error").Value = True
                                Rs.Fields.Item("SSN Description").Value = ErrorMsg & "Duplicate SSN - " & DUPSSN & " : "
                                Rs.Update()
                            End If
                            Rs.MoveNext()
                        Loop
                        Rs.Close()
                    End If
                    RsDupCheck.MoveNext()
                Loop
            End If

            'Housekeeping
            RsDupCheck.Close()
            RosterSSNConnection.Close()
            RosterConnection.Close()

            'Refreshes the columns that were just edited
            Form1.RosterDataGridView.Refresh()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

        Form1.Cursor = Cursors.Default

    End Sub

    Public Sub Roster_SSN_View()
        '***********************************************************************
        ' Sub to Change the Roster Grid View to just show SSN Errors From the Roster Table
        '***********************************************************************

        Dim ConnectionString As String
        Dim SqlStr As String
        Dim Connection As OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim Rs As New DataSet
        Dim x As Integer

        'Connect to the database
        ConnectionString = Client_Conn

        'Need to Connect to the Access Db
        Connection = New OleDbConnection(ConnectionString)
        Connection.Open()

        'SqlStr = "Select [ID], [EMPLOYEE ID], [EMPLOYEE LAST NAME], [EMPLOYEE FIRST NAME], [EMPLOYEE MIDDLE NAME], " &
        '" [EMPLOYEE DATE OF BIRTH], [EMPLOYEE SS#], [LOCATION NAME], [LOCATION NUMBER], [HIRE DATE], [TERMINATED DATE], " &
        '" [DATE ERROR], [DATE DESCRIPTION], [SSN ERROR], [SSN DESCRIPTION], [OTHER ERROR], [OTHER DESCRIPTION] " &
        '" FROM [ROSTER] WHERE [SSN ERROR] = '-1' ORDER BY [ID] ;"

        SqlStr = "SELECT * FROM [ROSTER] WHERE [SSN ERROR] = '-1' ORDER BY [ID] ;"
        oledbAdapter = New OleDbDataAdapter(SqlStr, ConnectionString)
        oledbAdapter.Fill(Rs)

        'Loads the Grid from the SQL Statement
        Form1.RosterDataGridView.DataSource = Rs.Tables(0)

        'Displays the Record Count
        x = Form1.RosterDataGridView.Rows.Count
        Form1.StatusLabelGridCount.Text = " Roster Records - " & Format(x, "#,###")

        'HouseKeeping
        Connection.Close()
        oledbAdapter.Dispose()

    End Sub

    Public Sub SSN_QC_Trans()
        '*************************************************
        ' Sub Routine to QC SSN numbers in the Transcription Table
        '*************************************************
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Dim TransSSN As New ADODB.Recordset
        Dim SSN, FormatSSN As String
        Dim Post As Integer

        Roster_Connection = Client_Conn

        Try
            FormatSSN = ""
            RosterConnection.Open(Roster_Connection)
            TransSSN.Open("SELECT [ID], [EMPLOYEE SS#] " &
            " FROM [I9] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If TransSSN.RecordCount = 0 Then Exit Sub
            If TransSSN.RecordCount <> 0 Then TransSSN.MoveFirst()

            'Loop through the Recordset
            Do While Not TransSSN.EOF
                SSN = If(IsDBNull(TransSSN.Fields.Item("EMPLOYEE SS#").Value), String.Empty, TransSSN.Fields.Item("EMPLOYEE SS#").Value)
                'Check for the correct format of the field
                If Len(SSN) = 11 And Post = InStr(4, SSN, "-") = 0 Then
                    'Has the right amount of characters and dashes
                    FormatSSN = SSN
                ElseIf Len(SSN) = 9 Then
                    'If the length is 9 and not less than need to place the - in the right place
                    FormatSSN = Left(SSN, 3) & "-" & Mid(SSN, 4, 2) & "-" & Right(SSN, 4)
                End If
                'Write the valid SSN Number Format the field
                TransSSN.Fields.Item("Employee SS#").Value = FormatSSN
                TransSSN.Update()
                FormatSSN = ""
                TransSSN.MoveNext()
            Loop

            'Housekeeping
            TransSSN.Close()
            RosterConnection.Close()

            'Refreshes the columns that were just edited
            Form1.I9DataGridView.Refresh()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Module
