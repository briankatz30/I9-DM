Module M_Date_QC
    Public Sub Roster_QC_Check_Dates()
        '**********************************************
        ' Sub Rountine to QC Dates from the Roster Table
        '**********************************************
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Dim RosterRs As New ADODB.Recordset
        Dim RosterFix As New ADODB.Recordset
        Dim Age As Integer
        Dim DOB As String
        Dim DateDesc As String
        Dim HDate As String
        Dim TermDate As String
        Dim ErrorMsg As String
        Dim TotalRow, CurrentRow As Integer

        Roster_Connection = Client_Conn

        Try
            ' Need to wipe the Date columns before we start
            RosterConnection.Open(Roster_Connection)
            RosterFix.Open("SELECT [Date Error], [Date Description] FROM [ROSTER] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            RosterFix.MoveFirst()
            CurrentRow = 1
            Do While Not RosterFix.EOF
                RosterFix.Fields.Item("DATE ERROR").Value = DBNull.Value.ToString
                RosterFix.Fields.Item("DATE DESCRIPTION").Value = DBNull.Value.ToString
                RosterFix.MoveNext()
            Loop

            RosterFix.Close()

            RosterRs.Open("SELECT [ID], [EMPLOYEE ID], [EMPLOYEE DATE OF BIRTH], [HIRE DATE], [TERMINATED DATE], [Date Error], [Date Description] " &
            " FROM [ROSTER] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RosterRs.MoveFirst()
            TotalRow = RosterRs.RecordCount
            Form1.TSProgressBar.Maximum = TotalRow
            Form1.TSProgressBar.Minimum = 1
            ' Loop through the Recordset
            Do While Not RosterRs.EOF
                '------------------------------------------ DOB -------------------------------
                'Need to check if Fields are Null
                DOB = If(IsDBNull(RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value), String.Empty, RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value)
                DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                HDate = If(IsDBNull(RosterRs.Fields.Item("HIRE DATE").Value), String.Empty, RosterRs.Fields.Item("HIRE DATE").Value)
                TermDate = If(IsDBNull(RosterRs.Fields.Item("TERMINATED DATE").Value), String.Empty, RosterRs.Fields.Item("TERMINATED DATE").Value)

                If DOB = "" Then
                    ' No Date Log the Error
                    If DateDesc = "" Then
                        ErrorMsg = ""
                    Else
                        ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                    End If
                    ' Write to the Description field
                    RosterRs.Fields.Item("DATE ERROR").Value = True
                    RosterRs.Fields.Item("DATE DESCRIPTION").Value = ErrorMsg & " DOB is missing - ID " & RosterRs.Fields.Item("ID").Value & " : "
                    RosterRs.Update()
                Else
                    ' We have data in the field Check and now see if it is a true date
                    If IsDate(DOB) Then
                        ' Caluclate the age for DOB under 16
                        Age = DateDiff("yyyy", CDate(DOB), Now())
                        If Age > 16 Then
                            ' Record is good on age write and format the date into the DOB field
                            RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value = Format(CDate(DOB), "M/d/yyyy")
                            RosterRs.Update()
                        Else
                            ' Log the Error
                            DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                            If DateDesc = "" Then
                                ErrorMsg = ""
                            Else
                                ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                            End If
                            ' Write the description
                            RosterRs.Fields.Item("DATE ERROR").Value = True
                            RosterRs.Fields.Item("DATE DESCRIPTION").Value = ErrorMsg & " Warning DOB is under Age (" & Age & ") : "
                            RosterRs.Update()
                        End If
                    Else
                        ' Log the Errror
                        DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                        If DateDesc = "" Then
                            ErrorMsg = ""
                        Else
                            ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                        End If
                        RosterRs.Fields.Item("DATE ERROR").Value = True
                        RosterRs.Fields.Item("DATE DESCRIPTION").Value = ErrorMsg & " DOB is not a valid date - " & DOB & " : "
                        RosterRs.Update()
                    End If
                End If
                '------------------------------------------ HIRE DATE -------------------------------
                ' Need to check if Hire Date is Null
                If HDate = "" Then
                    ' No Date Log the Error
                    DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                    If DateDesc = "" Then
                        ErrorMsg = ""
                    Else
                        ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                    End If
                    RosterRs.Fields.Item("Date Error").Value = True
                    RosterRs.Fields.Item("Date Description").Value = ErrorMsg & " Missing Hire Date " & " : "
                    RosterRs.Update()
                Else
                    ' We have data in the field Check to see if it is a true date
                    If IsDate(HDate) Then
                        ' Check to see if the Hire Date is greater than or equal to today's date
                        If CDate(HDate) > Now() Then
                            ' Log the Error
                            DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                            If DateDesc = "" Then
                                ErrorMsg = ""
                            Else
                                ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                            End If
                            RosterRs.Fields.Item("Date Error").Value = True
                            RosterRs.Fields.Item("Date Description").Value = ErrorMsg & " Hire Date > today -  " & HDate & " : "
                            RosterRs.Update()
                        Else
                            'Write and format the date into the Hire Date field
                            RosterRs.Fields.Item("Hire Date").Value = Format(CDate(HDate), "M/d/yyyy")
                            RosterRs.Update()
                        End If
                    Else
                        ' Log the errror in the record
                        DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                        If DateDesc = "" Then
                            ErrorMsg = ""
                        Else
                            ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                        End If

                        ' ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                        RosterRs.Fields.Item("Date Error").Value = True
                        RosterRs.Fields.Item("Date Description").Value = ErrorMsg & " Hire Date is invalid -  " & HDate & " : "
                        RosterRs.Update()

                    End If
                End If
                '------------------------------------------ TERMINATION DATE -------------------------------
                ' Need to check if Termination Date is Null - optional and not required
                If TermDate = "" Then
                    ' Null Date
                Else
                    'We have data in the field Check to see if it is a true date
                    If IsDate(CDate(TermDate)) And IsDate(CDate(HDate)) Then
                        'Check to see if the Termination Date is greater than or equal to Hire date
                        If CDate(CDate(TermDate)) < CDate(CDate(HDate)) Then
                            'Log the Error
                            DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                            If DateDesc = "" Then
                                ErrorMsg = ""
                            Else
                                ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                            End If
                            RosterRs.Fields.Item("Date Error").Value = True
                            RosterRs.Fields.Item("Date Description").Value = ErrorMsg & " Term Date greater than Hire Date  -  " & TermDate & " : "
                            RosterRs.Update()
                        Else
                        End If
                        ' Write and format the date into the Termination Date field
                        RosterRs.Fields.Item("Terminated Date").Value = Format(CDate(TermDate), "M/d/yyyy")
                        RosterRs.Update()
                    Else
                        ' Log the errror in the record
                        If IsDate(CDate(TermDate)) Or TermDate = "" Then
                            ' Good Date
                        Else
                            ' Bad Date
                            DateDesc = If(IsDBNull(RosterRs.Fields.Item("DATE DESCRIPTION").Value), String.Empty, RosterRs.Fields.Item("DATE DESCRIPTION").Value)
                            If DateDesc = "" Then
                                ErrorMsg = ""
                            Else
                                ErrorMsg = RosterRs.Fields.Item("DATE DESCRIPTION").Value
                            End If
                            RosterRs.Fields.Item("Date Error").Value = True
                            RosterRs.Fields.Item("Date Description").Value = ErrorMsg & " Term Date is invalid -  " & TermDate & " : "
                            RosterRs.Update()
                        End If
                    End If
                End If

                'Clears the Error Message
                ErrorMsg = ""
                RosterRs.MoveNext()
                Form1.TSProgressBar.Value = CurrentRow
                CurrentRow = CurrentRow + 1
            Loop

            'Refreshes the columns that were just edited
            Form1.RosterDataGridView.Refresh()

            RosterRs.Close()
            RosterConnection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Roster_Date_View()
        '***********************************************************************************
        ' Popup Menu Item to change the the Roster Grid View to just show Date Errors From the Roster Table
        '***********************************************************************************

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

        ''SqlStr = "Select [ID], [EMPLOYEE ID], [EMPLOYEE LAST NAME], [EMPLOYEE FIRST NAME], [EMPLOYEE MIDDLE NAME], " &
        '" [EMPLOYEE DATE OF BIRTH], [EMPLOYEE SS#], [LOCATION NAME], [LOCATION NUMBER], [HIRE DATE], [TERMINATED DATE], " &
        '" [DATE ERROR], [DATE DESCRIPTION], [SSN ERROR], [SSN DESCRIPTION], [OTHER ERROR], [OTHER DESCRIPTION] " &
        '" FROM [ROSTER] WHERE [DATE ERROR] = '-1' ORDER BY [ID] ;"

        SqlStr = "Select * FROM [ROSTER] WHERE [DATE ERROR] = '-1' ORDER BY [ID];"

        oledbAdapter = New OleDbDataAdapter(SqlStr, ConnectionString)
        oledbAdapter.Fill(Rs)

        'Loads the Grid with the SQL Results
        Form1.RosterDataGridView.DataSource = Rs.Tables(0)
        'Displays the Record Count
        x = Form1.RosterDataGridView.Rows.Count
        Form1.StatusLabelGridCount.Text = " Roster Records - " & Format(x, "#,###")

        'House Keeping
        Connection.Close()
        oledbAdapter.Dispose()

    End Sub

    Public Sub Date_QC_Trans()
        '*****************************************************************************
        ' Sub Routine to QC the Transcription Date of Birth and fix the format of that field for matching
        '*****************************************************************************
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Dim RosterRs As New ADODB.Recordset
        Dim DOB As String

        Roster_Connection = Client_Conn

        Try

            RosterConnection.Open(Roster_Connection)
            RosterRs.Open("SELECT [ID], [EMPLOYEE DATE OF BIRTH] FROM [ROSTER] ORDER BY [ID] ;",
            RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            RosterRs.MoveFirst()

            ' Loop through the Recordset loking at the DOB Field
            Do While Not RosterRs.EOF
                ' Need to check if Fields are Null
                DOB = If(IsDBNull(RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value), String.Empty, RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value)
                If DOB = "" Then
                    ' Nothing in the field move to next record
                Else
                    ' We have data in the field check to see if it is a true date
                    If IsDate(DOB) Then
                        RosterRs.Fields.Item("EMPLOYEE DATE OF BIRTH").Value = Format(CDate(DOB), "M/d/yyyy")
                        RosterRs.Update()
                    Else
                        ' Not a good date move to the next record
                    End If
                    ' Not a good date move to the next record
                End If
                RosterRs.MoveNext()
            Loop

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

        'HouseKeeping
        RosterRs.Close()
        RosterConnection.Close()

    End Sub

End Module
