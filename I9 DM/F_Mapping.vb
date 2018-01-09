Public Class F_Mapping


    Private Sub F_Mapping_Load(sender As Object, e As EventArgs) Handles Me.Load
        '*********************************************
        ' Setting when you open the Mapping Form  
        '*********************************************
        FieldMappingGV.RowHeadersDefaultCellStyle.Padding = New Padding(FieldMappingGV.RowHeadersWidth)
        FieldMappingGV.RowHeadersVisible = False

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '*********************************************
        ' Run Button to gather all the lines in the grid and create 
        ' and SQL statement to update the Roster Table with the
        ' Excel file that was provided.
        '*********************************************
        Dim SqlStmt As String
        Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        'Dim MyExcelDataQuery As String = "SELECT * FROM  [SHEET1$]"
        Dim Ds As New DataSet
        Dim Dt As New DataTable

        Try

            'Build the SQL Statement for the DataTable
            SqlStmt = "SELECT "
            For Each Row As DataGridViewRow In FieldMappingGV.Rows
                'Looks to see if the user checked the include box
                If Row.Cells(0).Value.ToString Then
                    SqlStmt = SqlStmt & "[" & Row.Cells(1).Value.ToString & "], "
                End If
            Next

            'Need to remove the last comma of the SQL statement
            SqlStmt = SqlStmt.Remove(SqlStmt.Length - 2)
            SqlStmt = SqlStmt & " FROM [SHEET1$] "

            Console.Write(SqlStmt)

            'Create a connection string to Excel
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & RosterExcelFilePath & " '; " & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;';")
            MyCommand = New OleDbDataAdapter(SqlStmt, MyConnection)
            MyCommand.Fill(Ds)
            Dt = Ds.Tables(0)
            For Each Row As DataRow In Dt.Rows
                Dt.Rows(0).ToString()
                Console.Write(Row.Field(Of String)(1).ToString)
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub


End Class