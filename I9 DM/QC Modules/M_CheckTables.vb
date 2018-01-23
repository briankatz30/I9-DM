
Imports System.Data.SqlClient

Module M_CheckTables
    '********************************************************************************************
    '  Module to load All the DataGridViews from SQLDb.  Also runs a SQL Stored Procedure to update the I9 Table
    ' rows that match records from the roster table to the transcript tables by SSN, First, Last and DOB
    '*********************************************************************************************
    Public Sub Load_GridViews()
        '****************************************************
        '  Public Sub Routine to load all the GridViews on the Main Form
        '****************************************************
        Dim oledbAdapter As OleDbDataAdapter
        Dim I9S As New DataSet
        Dim Rs As New DataSet
        Dim NotMatch As New DataSet
        Dim SSNMatch As New DataSet
        Dim RsCount As New ADODB.Recordset
        Dim RsCon As New ADODB.Connection
        Dim TotalRosterCount, TotalTransCount As Integer

        'Resets the Datasource for each DataGridView to Nothing for viewing purposes
        'Form1.RosterDataGridView.DataSource = Nothing
        Form1.I9DataGridView.DataSource = Nothing
        Form1.MatchDataGridView.DataSource = Nothing
        Form1.DGVNotMatchTrans.DataSource = Nothing
        Form1.DGVRosterMatch.DataSource = Nothing

        Try
            'Load the Roster Grid View with the Roster Table
            Form1.ToolStripStatusLabel2.Text = "Loading Roster Data...."
            oledbAdapter = New OleDbDataAdapter("EXEC dbo.SP_LOAD_ROSTER", Client_Conn)
            oledbAdapter.Fill(Rs)
            Form1.RosterDataGridView.DataSource = Rs.Tables(0)
            Form1.RosterDataGridView.Refresh()

            'Load the Transcription Grid View with the I9 Table
            Form1.ToolStripStatusLabel2.Text = "Loading Transcription Data..."
            oledbAdapter = New OleDbDataAdapter("EXEC dbo.SP_LOAD_I9", Client_Conn)
            oledbAdapter.Fill(I9S)
            Form1.I9DataGridView.DataSource = I9S.Tables(0)
            Form1.I9DataGridView.Refresh()

            'Hides 2 of the Columns that are not needed
            Form1.I9DataGridView.Columns("2 Page Flag").Visible = False
            Form1.I9DataGridView.Columns("Section 3 Flag").Visible = False

            'Loads Records that have no match between the roster and I9 Table
            Form1.ToolStripStatusLabel2.Text = "Loading Non Matching Data...."
            oledbAdapter = New OleDbDataAdapter("EXEC [dbo].SP_I9_MATCHISNULL", Client_Conn)
            oledbAdapter.Fill(NotMatch)

            'Loads the No Match DataGridView
            Form1.DGVNotMatchTrans.DataSource = NotMatch.Tables(0)
            Form1.DGVNotMatchTrans.Refresh()
            Form1.NotMatchingTransTxt.Text = (Form1.DGVNotMatchTrans.RowCount - 1).ToString

            'Hides the Columns that are not needed
            Form1.DGVNotMatchTrans.Columns("2 Page Flag").Visible = False
            Form1.DGVNotMatchTrans.Columns("Section 3 Flag").Visible = False

            'Gets all Roster Records that do not match a transcription record by SSN or Last, First, Middle and DOB to not match
            'loads them into the RosterMatch DataGridView
            oledbAdapter = New OleDbDataAdapter("EXEC dbo.SP_ROSTER_MATCH", Client_Conn)
            oledbAdapter.Fill(SSNMatch)
            Form1.DGVRosterMatch.DataSource = SSNMatch.Tables(0)
            RosterNotMatchCount = (Form1.DGVRosterMatch.RowCount - 1).ToString
            Form1.TotalNotMatchRosterTxt.Text = RosterNotMatchCount

            TotalTransCount = Form1.I9DataGridView.Rows.Count - 1
            RosterMatch = Form1.DGVRosterMatch.Rows.Count - 1

            Form1.ToolStripStatusLabel2.Text = "Getting Roster Statistics.."

            'Get the Stats from the Roster Table and Loads them into the Roster ListView
            Get_Stats_Roster()

            'Builds the Transcription List View
            Form1.ToolStripStatusLabel2.Text = "Getting Transcription Statistics..."
            LoadTransLV()

            Dim i9Connection As New OleDbConnection(Client_Conn)
            i9Connection.Open()

            Dim cmd As OleDbCommand = New OleDbCommand("SP_I9VIEW", i9Connection)
            cmd.CommandType = CommandType.StoredProcedure
            Dim I9SSNCOUNT As OleDbParameter = cmd.Parameters.Add("@I9SSNCOUNT", OleDbType.Integer)
            I9SSNCOUNT.Direction = ParameterDirection.Output
            Dim I9DOBCOUNT As OleDbParameter = cmd.Parameters.Add("@I9DOBCOUNT", OleDbType.Integer)
            I9DOBCOUNT.Direction = ParameterDirection.Output
            Dim I9FLMICOUNT As OleDbParameter = cmd.Parameters.Add("@I9FIRSTLASTMIDDLECOUNT", OleDbType.Integer)
            I9FLMICOUNT.Direction = ParameterDirection.Output
            Dim I9NAMECOUNT As OleDbParameter = cmd.Parameters.Add("@I9NAMECOUNT", OleDbType.Integer)
            I9NAMECOUNT.Direction = ParameterDirection.Output
            cmd.ExecuteNonQuery()
            i9Connection.Close()

            Form1.MatchSSNTxt.Text = I9SSNCOUNT.Value
            Form1.MatchDOBTxt.Text = I9DOBCOUNT.Value
            Form1.MatchNameMItxt.Text = I9FLMICOUNT.Value
            Form1.MatchNametxt.Text = I9NAMECOUNT.Value

            'Displays the Record Count on the Status Bar - Roster
            TotalRosterCount = Form1.RosterDataGridView.Rows.Count - 1
            If TotalRosterCount > 0 Then
                Form1.StatusLabelGridCount.Text = " Roster Records - " & Format(TotalRosterCount, "#,###")
            Else
                'No Records Found
                Form1.StatusLabelGridCount.Text = " No Roster Records "
            End If

            'Displays the Record Count - Transaction
            TotalTransCount = Form1.I9DataGridView.Rows.Count - 1
            If TotalTransCount > 0 Then
                Form1.StatusLabelTransCount.Text = " Trans Records - " & Format(TotalTransCount, "#,###")
            Else
                'No records Found
                Form1.StatusLabelTransCount.Text = " No Trans Records "
            End If

            Form1.ToolStripStatusLabel2.Text = "Completed...."

            'Housekeeping
            oledbAdapter.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    Public Sub Match_Check()
        '**********************************************
        ' Sub Routine to Update the I9 table with the Match field where
        ' the SSN number in the Roster is the same as the I9 table
        '**********************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = "dbo.SP_RUN_MATCH_CHECK"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Module