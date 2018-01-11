Imports System.ComponentModel

Public Class F_Mapping

    Private Sub F_Mapping_Load(sender As Object, e As EventArgs) Handles Me.Load
        '*********************************************
        ' Setting when you open the Mapping Form  
        '*********************************************
        FieldMappingGV.RowHeadersDefaultCellStyle.Padding = New Padding(FieldMappingGV.RowHeadersWidth)
        FieldMappingGV.RowHeadersVisible = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        '*********************************************
        ' Run Button to gather all the lines in the grid and create 
        ' and SQL statement to update the Roster Table with the
        ' Excel file that was provided.
        '*********************************************
        Dim SqlStmt As String = Nothing
        Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim oledbAdapter As OleDbDataAdapter
        Dim Ds As New DataSet
        Dim DTable As New DataTable
        Dim Rs As New DataSet
        Dim FieldNameList As New List(Of String)()
        Dim i As Integer = 0
        Dim x As Integer
        Dim c As Integer = 0
        Dim TestName As String = Nothing
        Dim FieldNames As String = Nothing
        Dim PrimaryKey, ForeignKey As String

        Try

            'Need to add the first Column before we loop through
            'the user selections to ge the Primary and Foreign Keys
            ForeignKey = "[ROSTER].[" & ForeginKeyComboBox.Text & "]"
            FieldNameList.Add(PrimaryKeyComboBox.Text)

            'Build the SQL Statement for the DataTable
            For Each Row As DataGridViewRow In FieldMappingGV.Rows
                'Looks to see if the user checked the overlay box
                If Row.Cells(0).Value.ToString Then
                    SqlStmt = SqlStmt & "[" & Row.Cells(1).Value.ToString & "], "
                    FieldNameList.Add("[" & Row.Cells(3).Value.ToString & "]")
                End If
            Next

            'Adds the Primary Key from the combobox from the Spreadsheet
            'First column will always be the primary Key Column
            SqlStmt = " SELECT [" & PrimaryKeyComboBox.Text & "], " & SqlStmt

            'Need to remove the last comma of the SQL statement
            SqlStmt = SqlStmt.Remove(SqlStmt.Length - 2)
            SqlStmt = SqlStmt & " FROM [SHEET1$] "

            'Console.WriteLine(SqlStmt)

            'Create a connection string to Excel
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & RosterExcelFilePath & " '; " & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;';")
            MyCommand = New OleDbDataAdapter(SqlStmt, MyConnection)
            MyCommand.Fill(Ds)
            DTable = Ds.Tables(0)

            'Get the number of Columns that were selected
            x = DTable.Columns.Count - 1

            For Each Row As DataRow In DTable.Rows
                PrimaryKey = DTable.Rows(c).Item(0)
                For i = 1 To x
                    'All the Other Columns to Build the SQL Statement
                    FieldNames = FieldNames & FieldNameList.Item(i) & " = '" & DTable.Rows(c).Item(i) & "', "
                Next
                'String work on the SQL Statment
                FieldNames = FieldNames.Remove(FieldNames.Length - 2)
                FieldNames = "UPDATE ROSTER SET " & FieldNames & " WHERE " & ForeignKey & " = '" & PrimaryKey & "' ;"
                Console.WriteLine(FieldNames)
                'Need to execute the SQL Statement to update the row
                Update_Roster(FieldNames)
                'Reset the variable for the next loop
                FieldNames = Nothing
                c += 1
            Next

            'Refreshes the Grid view with the newly edited records
            Form1.ToolStripStatusLabel2.Text = "Loading Roster Data...."
            oledbAdapter = New OleDbDataAdapter("EXEC dbo.SP_LOAD_ROSTER", Client_Conn)
            oledbAdapter.Fill(Rs)
            Form1.RosterDataGridView.DataSource = Rs.Tables(0)
            Form1.RosterDataGridView.Refresh()
            Rs.Dispose()
            oledbAdapter.Dispose()

            MsgBox("Roster Updated", vbInformation)

            If FieldMappingGV.Columns.Contains("RosterColumnField") Then
                FieldMappingGV.Columns.Remove("RosterColumnField")
            End If

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub Update_Roster(ByVal SQLStatement As String)
        '**********************************************
        ' Sub Routine to Update the Roster Table from the Overlay
        ' Spreadsheet
        '**********************************************
        Dim Connection As New OleDbConnection(Client_Conn)
        Dim cmd As New OleDbCommand
        Dim rowsAffected As Integer

        Try
            Connection = New OleDbConnection(Client_Conn)
            Connection.Open()
            cmd.CommandText = SQLStatement
            cmd.CommandType = CommandType.Text
            cmd.Connection = Connection
            rowsAffected = cmd.ExecuteNonQuery()
            Connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        If FieldMappingGV.Columns.Contains("RosterColumnField") Then
            FieldMappingGV.Columns.Remove("RosterColumnField")
        End If
        Me.Close()

    End Sub

    Private Sub F_Mapping_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If FieldMappingGV.Columns.Contains("RosterColumnField") Then
            FieldMappingGV.Columns.Remove("RosterColumnField")
        End If

    End Sub

End Class