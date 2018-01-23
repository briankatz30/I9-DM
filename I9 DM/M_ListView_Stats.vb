
Module M_ListView_Stats

    Dim BadDateCount, BadSSNCount, BadOtherCount, AllBadCount, BadTotalRosterCount As Integer
    Dim TotalRosterCount As Integer
    Dim TotalTransCount As Integer
    Dim SQLStmt As String
    Dim PrectDate, PrectSSN, PrectOther, MeetPrect As Double
    Dim DoesnotMeetReq, TotalRosterMatch As Double
    Dim BadRosterMatch As Double

    Public Sub Get_Stats_Roster()
        '*******************************************************
        ' Sub Routine to load the Stats \ Errors of the Roster table on to 
        ' the F_Main Form for the ListView
        '*******************************************************

        'Dim Rs As New ADODB.Recordset
        Dim DbConnection As New OleDbConnection(Client_Conn)
        Dim PrectGoodDate, SSNGoodPrect, GoodOtherPrect As Double

        Try
            ' Create the command with the stored procedure 
            ' and add the parameters required'
            DbConnection.Open()
            Dim cmd1 As OleDbCommand = New OleDbCommand("SP_ROSTERVIEW", DbConnection)
            cmd1.CommandType = CommandType.StoredProcedure
            Dim SPRosterCount As OleDbParameter = cmd1.Parameters.Add("@ROSTERCOUNT", OleDbType.Integer)
            SPRosterCount.Direction = ParameterDirection.Output
            Dim SPI9Count As OleDbParameter = cmd1.Parameters.Add("I9COUNT", OleDbType.Integer)
            SPI9Count.Direction = ParameterDirection.Output
            Dim SPRosterDate As OleDbParameter = cmd1.Parameters.Add("ROSTERDATEERROR", OleDbType.Integer)
            SPRosterDate.Direction = ParameterDirection.Output
            Dim SPSSNRoster As OleDbParameter = cmd1.Parameters.Add("ROSTERSSNERROR", OleDbType.Integer)
            SPSSNRoster.Direction = ParameterDirection.Output
            Dim SPRosterOther As OleDbParameter = cmd1.Parameters.Add("ROSTEROTHERERRORS", OleDbType.Integer)
            SPRosterOther.Direction = ParameterDirection.Output
            Dim SPAll As OleDbParameter = cmd1.Parameters.Add("ROSTERALLERRORS", OleDbType.Integer)
            SPAll.Direction = ParameterDirection.Output

            cmd1.ExecuteNonQuery()

            ' Result of the Stored Procedure
            TotalRosterCount = SPRosterCount.Value
            BadDateCount = SPRosterDate.Value
            BadSSNCount = SPSSNRoster.Value
            BadOtherCount = SPRosterOther.Value
            AllBadCount = SPAll.Value
            TotalTransCount = SPI9Count.Value

            DbConnection.Close()

            ' Check to see if we have any records in the Roster Table
            If TotalRosterCount = 0 Then
                'No records
                BadDateCount = 0
                BadSSNCount = 0
                BadOtherCount = 0
                TotalRosterCount = 0
                DoesnotMeetReq = 0
                Exit Sub
            Else

                ' Check and load the totals from the Stored Procedure
                If BadDateCount = 0 Then
                    Form1.RosterDateCountTxt.Text = "0".ToString
                Else
                    Form1.RosterDateCountTxt.Text = Format(BadDateCount, "#,###").ToString
                End If

                If BadSSNCount = 0 Then
                    Form1.RosterSSNCountTxt.Text = "0".ToString
                Else
                    Form1.RosterSSNCountTxt.Text = Format(BadSSNCount, "#,###").ToString
                End If
                If BadOtherCount = 0 Then
                    Form1.RosterOtherCountTxt.Text = "0".ToString
                Else
                    Form1.RosterOtherCountTxt.Text = Format(BadOtherCount, "#,###").ToString
                End If
                If AllBadCount = 0 Then
                    Form1.MeetReqCountTxt.Text = Format(TotalRosterCount, "#,###").ToString & "  |  " & "0".ToString
                Else
                    If (TotalRosterCount - AllBadCount) = 0 Then
                        Form1.MeetReqCountTxt.Text = "0".ToString & "  |  " & Format(AllBadCount, "#,###").ToString
                    Else
                        Form1.MeetReqCountTxt.Text = Format((TotalRosterCount - AllBadCount), "#,###").ToString & "  |  " & Format(AllBadCount, "#,###").ToString
                    End If
                End If
            End If

            'ROSTER LIST VIEW
            Form1.RosterLV.Clear()
            Form1.RosterLV.GridLines = True
            ' Adds the Columns to the ListView
            Form1.RosterLV.Columns.Add("Items", 220, HorizontalAlignment.Left)
            Form1.RosterLV.Columns.Add("% Good", 110, HorizontalAlignment.Center)
            Form1.RosterLV.Columns.Add("% Error", 110, HorizontalAlignment.Center)
            Form1.RosterLV.Columns.Add("Total Records", 110, HorizontalAlignment.Center)

            'DISPLAY THE STATUS IN THE LISTVIEW ON THE STATS TAB
            'DATE
            PrectDate = (BadDateCount / TotalRosterCount)
            PrectGoodDate = ((TotalRosterCount - BadDateCount) / TotalRosterCount)
            Form1.RosterDatePrectxt.Text = Format(PrectDate, "0.00%").ToString
            Form1.RosterLV.Items.Add("Date")
            Form1.RosterLV.Items.Item(0).SubItems.Add(Format(PrectGoodDate, "0.00%").ToString & " (" & TotalRosterCount - BadDateCount & ")")
            Form1.RosterLV.Items.Item(0).SubItems.Add(Format(PrectDate, "0.00%").ToString & " (" & BadDateCount & ")")
            Form1.RosterLV.Items.Item(0).SubItems.Add(Format(TotalRosterCount, "#,###").ToString)

            'SSN
            PrectSSN = (BadSSNCount / TotalRosterCount)
            SSNGoodPrect = ((TotalRosterCount - BadSSNCount) / TotalRosterCount)
            Form1.RosterSSNPrectTxt.Text = Format(PrectSSN, "0.00%").ToString
            Form1.RosterLV.Items.Add("SSN")
            Form1.RosterLV.Items.Item(1).SubItems.Add(Format(SSNGoodPrect, "0.00%").ToString & " (" & TotalRosterCount - BadSSNCount & ")")
            Form1.RosterLV.Items.Item(1).SubItems.Add(Format(PrectSSN, "0.00%").ToString & " (" & BadSSNCount & ")")

            'OTHER
            PrectOther = (BadOtherCount / TotalRosterCount)
            GoodOtherPrect = ((TotalRosterCount - BadOtherCount) / TotalRosterCount)
            Form1.RosterOtherPrectTxt.Text = Format(PrectOther, "0.00%").ToString
            Form1.RosterLV.Items.Add("Other")
            Form1.RosterLV.Items.Item(2).SubItems.Add(Format(GoodOtherPrect, "0.00%").ToString & " (" & TotalRosterCount - BadOtherCount & ")")
            Form1.RosterLV.Items.Item(2).SubItems.Add(Format(PrectOther, "0.00%").ToString & " (" & BadOtherCount & ")")

            'MEETS REQUIREMENTS
            DoesnotMeetReq = (AllBadCount / TotalRosterCount)
            MeetPrect = ((TotalRosterCount - AllBadCount) / TotalRosterCount)
            Form1.RosterLV.Items.Add("Meeting Requirements")
            Form1.RosterLV.Items.Item(3).SubItems.Add(Format(MeetPrect, "0.00%").ToString & " (" & TotalRosterCount - AllBadCount & ")")
            Form1.RosterLV.Items.Item(3).SubItems.Add(Format(DoesnotMeetReq, "0.00%").ToString & " (" & AllBadCount & ")")

            'ROSTER RECORDS
            BadTotalRosterCount = TotalRosterCount - RosterNotMatchCount
            Form1.RosterLV.Items.Add("")
            Form1.RosterLV.Items.Add("Roster not MatchingTranscription")
            TotalRosterMatch = (RosterMatch / TotalRosterCount)
            BadRosterMatch = ((TotalRosterCount - RosterNotMatchCount) / TotalRosterCount)
            Form1.RosterLV.Items.Item(5).SubItems.Add(Format(BadRosterMatch, "0.00%").ToString & " (" & BadTotalRosterCount & ")")
            Form1.RosterLV.Items.Item(5).SubItems.Add(Format(TotalRosterMatch, "0.00%").ToString & " (" & RosterNotMatchCount & ")")
            Form1.RosterLV.Items.Item(5).SubItems.Add(Format(TotalRosterCount, "#,###").ToString)

            'DISPLAY NEXT ON THE ROSTER TAB
            If (TotalRosterCount - AllBadCount) < 1 Then
                Form1.RosterMeetReqPrectTxt.Text = "0".ToString & "  |  " & Format(DoesnotMeetReq, "0.00%").ToString
            Else
                Form1.RosterMeetReqPrectTxt.Text = Format(MeetPrect, "0.00%").ToString & "  |  " & Format(DoesnotMeetReq, "0.00%").ToString
            End If

        Catch ex As Exception
            'MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Clear_Stats()

        '**********************************************
        ' Sub to Clear the Stats of the table on to the F_Main Form
        '**********************************************

        'Update the Text boxes for the Stats
        'Date Fields
        'Forms!F_Main.BadDateCtTxt.SetFocus
        'Forms!F_Main.BadDateCtTxt.Text = ""
        'Forms!F_Main.ChkMakDate.Visible = False
        'Forms!F_Main.DatePrectxt.Visible = False


        'SSN Fields
        'Forms!F_Main.BadSSNCtTxt.SetFocus
        'Forms!F_Main.BadSSNCtTxt.Text = ""
        'Forms!F_Main.ChkMakSSN.Visible = False
        'Forms!F_Main.SSNPrectxt.Visible = False

        'Other Fields
        'Forms!F_Main.BadOtherCtTxt.SetFocus
        'Forms!F_Main.BadOtherCtTxt.Text = ""
        'Forms!F_Main.ChkMakOther.Visible = False
        'Forms!F_Main.OtherPrectxt.Visible = False

        'Meet Requirements
        'Forms!F_Main.MeetReqCtTxt.SetFocus
        'Forms!F_Main.MeetReqCtTxt.Text = ""
        'Forms!F_Main.ChkMakMeet.Visible = False
        'Forms!F_Main.GoodPrectxt.Visible = False

        'Total Employee Count
        'Forms!F_Main.RosterCountTxt.Visible = False
        'Forms!F_Main.RCmdOpenProject.SetFocus

        'Status Bar
        'SysCmd acSysCmdSetStatus, "No Project Loaded "

    End Sub

    Public Sub LoadTransLV()

        '*****************************************************
        ' Sub Routine to Load the Stats from the transcription Table in to 
        ' the Transcription List View
        '*****************************************************
        'Dim RosterTransConnection As New ADODB.Connection
        Dim Rs As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Try
            Roster_Connection = Client_Conn

            RosterConnection.Open(Roster_Connection)
            SQLStmt = "SELECT COUNT(*) FROM [I9];"
            Rs.Open(SQLStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.MoveFirst()
            If Rs.EOF = True Or IsNothing(Rs.EOF) Then
                TotalTransCount = 0
            Else
                TotalTransCount = Rs.Fields.Item(0).Value
            End If

            'Transcription List View
            Form1.TransLV.Clear()
            Form1.TransLV.GridLines = True


            'Adds the Columns to the ListView
            Form1.TransLV.Columns.Add("Items", 220, HorizontalAlignment.Left)
            Form1.TransLV.Columns.Add(" Audit ", 110, HorizontalAlignment.Center)
            Form1.TransLV.Columns.Add("Warnings", 110, HorizontalAlignment.Center)
            Form1.TransLV.Columns.Add(" Total Records ", 110, HorizontalAlignment.Center)

            'Add Data to the List View
            Form1.TransLV.Items.Add("SSN Numbers")
            Form1.TransLV.Items.Add("DOBs")
            Form1.TransLV.Items.Add("Orphan Recs")
            Form1.TransLV.Items.Add("Form Version")
            Form1.TransLV.Items.Add("Section 3s")
            Form1.TransLV.Items.Add("Signature")
            Form1.TransLV.Items.Add("")
            Form1.TransLV.Items.Add("Total Transcription Records")
            Form1.TransLV.Items.Item(7).SubItems.Add("")
            Form1.TransLV.Items.Item(7).SubItems.Add("")
            Form1.TransLV.Items.Item(7).SubItems.Add(Format(TotalTransCount, "#,###"))

            Rs.Close()
            RosterConnection.Close()

            'Loads the SSN Counts in the List View
            Get_Transciption_Stats()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Get_Transciption_Stats()
        '*************************************************
        ' Sub Routine to QC SSN numbers in the Transcription Table
        '*************************************************
        Dim RosterConnection As New ADODB.Connection
        Dim TransSSN As New ADODB.Recordset
        Dim OrphanRs As New ADODB.Recordset
        Dim Section3Rs As New ADODB.Recordset
        Dim Post, nPost As Integer
        Dim GoodCountSSN As Integer = 0
        Dim BadCountSSN As Integer = 0
        Dim BadDOBCount As Integer = 0
        Dim GoodDOBCount As Integer = 0
        Dim FormCount As Integer = 0
        Dim MissingForm As Integer = 0
        Dim OrphanCount As Integer = 0
        Dim NoOrphanCount As Integer = 0
        Dim Section3Count As Integer = 0
        Dim NoSection3Count As Integer = 0
        Dim DOB, SqlStmt, SSN, FormatSSN, FormVersion As String

        Try

            FormatSSN = ""
            RosterConnection.Open(Client_Conn)
            TransSSN.Open("SELECT [ID], [EMPLOYEE SS#], [EMPLOYEE DATE OF BIRTH], [FORM VERSION], [ORPHANDOC], [2 PAGE FLAG] " &
            " FROM [I9] ORDER BY [ID] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If TransSSN.RecordCount = 0 Then Exit Sub
            If TransSSN.RecordCount <> 0 Then TransSSN.MoveFirst()

            'Loop through the Recordset
            'checking for SSN Number
            Do While Not TransSSN.EOF
                SSN = If(IsDBNull(TransSSN.Fields.Item("EMPLOYEE SS#").Value), String.Empty, TransSSN.Fields.Item("EMPLOYEE SS#").Value)
                'Check for the correct format of the field
                If Len(SSN) = 11 And Post = InStr(4, SSN, "-") = 0 And nPost = InStr(7, SSN, "-") = 0 Then
                    GoodCountSSN = GoodCountSSN + 1
                Else
                    BadCountSSN = BadCountSSN + 1
                End If
                TransSSN.MoveNext()
            Loop

            'Date of Birth Check
            TransSSN.MoveFirst()
            'Loop through the Recordset
            Do While Not TransSSN.EOF
                DOB = If(IsDBNull(TransSSN.Fields.Item("Employee Date of Birth").Value), String.Empty, TransSSN.Fields.Item("Employee Date of Birth").Value)
                If IsDate(DOB) Then
                    GoodDOBCount = GoodDOBCount + 1
                Else
                    BadDOBCount = BadDOBCount + 1
                End If
                TransSSN.MoveNext()
            Loop

            'Check for Form Versions
            TransSSN.MoveFirst()
            'Loop through the Recordset
            Do While Not TransSSN.EOF
                FormVersion = If(IsDBNull(TransSSN.Fields.Item("Form Version").Value), String.Empty, TransSSN.Fields.Item("Form Version").Value)
                If Len(FormVersion) > 0 Then
                    FormCount = FormCount + 1
                Else
                    MissingForm = MissingForm + 1
                End If
                TransSSN.MoveNext()
            Loop

            'Close the Recordset
            TransSSN.Close()

            'Orphan Record Count
            SqlStmt = "SELECT COUNT([ORPHANDOC]) FROM [I9] WHERE [ORPHANDOC] = 'Y'; "
            OrphanRs.Open(SqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            OrphanCount = OrphanRs(0).Value

            'Close the Recordset
            OrphanRs.Close()

            SqlStmt = "SELECT COUNT([ORPHANDOC]) FROM [I9] WHERE [ORPHANDOC] IS NULL OR  [ORPHANDOC]= ''; "
            OrphanRs.Open(SqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            NoOrphanCount = OrphanRs(0).Value

            'Close the Recordset
            OrphanRs.Close()

            'Section 3 Record Count
            SqlStmt = "SELECT COUNT([SECTION 3 FLAG]) FROM [I9] WHERE [SECTION 3 FLAG] = '1'; "
            Section3Rs.Open(SqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            Section3Count = Section3Rs(0).Value

            'Close the Recordset
            Section3Rs.Close()

            SqlStmt = "SELECT COUNT([SECTION 3 FLAG]) FROM [I9] WHERE [SECTION 3 FLAG] IS NULL OR [SECTION 3 FLAG] = '' ; "
            Section3Rs.Open(SqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            NoSection3Count = Section3Rs(0).Value

            'Close the Recordset
            Section3Rs.Close()

            'Update Listview
            Form1.TransLV.Items.Item(0).UseItemStyleForSubItems = False
            Form1.TransLV.Items.Item(1).UseItemStyleForSubItems = False
            Form1.TransLV.Items.Item(2).UseItemStyleForSubItems = False
            Form1.TransLV.Items.Item(3).UseItemStyleForSubItems = False
            Form1.TransLV.Items.Item(4).UseItemStyleForSubItems = False

            Form1.TransLV.Items.Item(0).SubItems.Add(Format(GoodCountSSN, "#,###")).ForeColor = Color.DarkCyan
            Form1.TransLV.Items.Item(0).SubItems.Add(Format(BadCountSSN, "#,###")).ForeColor = Color.Red
            Form1.TransLV.Items.Item(1).SubItems.Add(Format(GoodDOBCount, "#,###")).ForeColor = Color.DarkCyan
            Form1.TransLV.Items.Item(1).SubItems.Add(Format(BadDOBCount, "#,###")).ForeColor = Color.Red
            Form1.TransLV.Items.Item(2).SubItems.Add(Format(OrphanCount, "#,###")).ForeColor = Color.DarkCyan
            Form1.TransLV.Items.Item(2).SubItems.Add(Format(NoOrphanCount, "#,###")).ForeColor = Color.Red
            Form1.TransLV.Items.Item(3).SubItems.Add(Format(FormCount, "#,###")).ForeColor = Color.DarkCyan
            Form1.TransLV.Items.Item(3).SubItems.Add(Format(MissingForm, "#,###")).ForeColor = Color.Red
            Form1.TransLV.Items.Item(4).SubItems.Add(Format(NoSection3Count, "#,###")).ForeColor = Color.DarkCyan
            Form1.TransLV.Items.Item(4).SubItems.Add(Format(Section3Count, "#,###")).ForeColor = Color.Red

            'Housekeeping
            RosterConnection.Close()

            'Refreshes the columns that were just edited in the transcript table
            Form1.I9DataGridView.Refresh()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
End Module
