﻿
Imports System.Data.SqlClient

Module M_Import_Excel

    Public Sub ImportDataFromExcel(ExcelFilePath As String)
        '*************************************************************
        ' Module to import an Excel Overlay Spreadsheet for receiving client 
        ' data to fix errors and allowing the user to mapped the fields to the roster
        ' table
        '*************************************************************

        Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim MyExcelDataQuery As String = "SELECT * FROM  [SHEET1$]"
        Dim Ds As New DataSet
        Dim Dt As New DataTable
        Dim Rs As New DataSet
        Dim Rt As New DataTable
        Dim Conn As New SqlConnection

        Try
            'Create a connection string to Excel
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & ExcelFilePath & " '; " & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;';")
            MyCommand = New OleDbDataAdapter(MyExcelDataQuery, MyConnection)
            MyCommand.Fill(Ds)
            Dt = Ds.Tables(0)

            'Need to load the combobox in the grid with the field names of the Roster Table
            Dim cmb As New DataGridViewComboBoxColumn()
            cmb.HeaderText = "Roster Field Names"
            cmb.Name = "RosterColumnField"
            MyCommand = New OleDbDataAdapter("Select Name FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Roster') and column_id Between 2 and 36 ;", Client_Conn)
            MyCommand.Fill(Rs)
            Rt = Rs.Tables(0)

            'Loads the Roster Columns into the Grid
            F_Mapping.ForeginKeyComboBox.Items.Add("ID")
            cmb.Items.Add(" ")
            For Each Row As DataRow In Rt.Rows
                cmb.Items.Add(Row(0).ToString)
                F_Mapping.ForeginKeyComboBox.Items.Add(Row(0).ToString)
            Next

            'Add the Combobox to the grid
            F_Mapping.FieldMappingGV.Columns.Add(cmb)
            F_Mapping.FieldMappingGV.Columns(3).Width = 300

            'Close the adaptor
            MyCommand.Dispose()

            'Load the field header row from the spreadsheet to the grid
            For Each Column As DataColumn In Dt.Columns
                F_Mapping.FieldMappingGV.Rows.Add(False, Column.ColumnName, "->")
                F_Mapping.PrimaryKeyComboBox.Items.Add(Column.ColumnName)
            Next

            Dt.Dispose()
            Rt.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Function Locate_Field(ByVal FieldName As String) As Boolean

        Dim SqlStmt As String
        Dim Reader As SqlDataReader
        Dim Conn As New SqlConnection
        Dim SqlConnect As String

        SqlConnect = Replace(Client_Conn, "Provider = SQLOLEDB.1;", "")
        Conn = New SqlConnection(SqlConnect)
        Conn.Open()
        'Sql Stmt to locate the field name inside the Roster Table
        SqlStmt = "SELECT Name FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Roster') and (column_id Between 2 and 36) and Name = '" & FieldName & "'; "
        Dim cmd As New SqlCommand(SqlStmt, Conn)
            Reader = cmd.ExecuteReader
            If Reader.HasRows Then
                Return True
            Else
                Return False
            End If

    End Function

End Module
