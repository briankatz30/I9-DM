Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Module M_ExportExcel
    'Path for getting the file that you need to get the empty spreadsheet
    Dim resourcesFolder = IO.Path.GetFullPath(Application.StartupPath & "\Resources\")

    'Use the Empty Spreadsheet as a template
    Dim fileName = "Empty.xlsx"

    Public Sub Export_Excel(DGV As DataGridView)
        '**************************************
        '  Public Sub to Export the Grid view to Excel
        '*************************************
        Dim xlApp As New Excel.Application
        Dim Worksheet As Excel.Worksheet
        Dim Workbook As Excel.Workbook
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        Dim Proc As System.Diagnostics.Process
        Dim misValue As Object = System.Reflection.Missing.Value

        Try
            Form1.Cursor = Cursors.WaitCursor
            'Test to make sure Excel is installed correctly
            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
                Exit Sub
            End If

            'Open the spreadsheet file
            'Workbook = xlApp.Workbooks.Open(resourcesFolder & fileName)
            Workbook = xlApp.Workbooks.Add(misValue)
            Worksheet = Workbook.Worksheets("Sheet1")

            ' Loop thru the Column headers and write to Excel
            For Each dc In DGV.Columns
                colIndex = colIndex + 1
                Worksheet.Cells(1, colIndex) = dc.Name
            Next

            ' Loop thru the Rows of the Grid and write to Excel
            For i As Integer = 0 To DGV.Rows.Count - 2 Step +1
                For j As Integer = 0 To DGV.Columns.Count - 1 Step +1
                    Worksheet.Cells(i + 2, j + 1).Value = DGV.Item(j, i).Value.ToString
                    Form1.ToolStripStatusLabel2.Text = "Loading Records into Excel " & i & " of  " & DGV.Rows.Count - 2
                Next
            Next

            ' Save the Excel file to a user location
            Using SFD As New SaveFileDialog
                If SFD.ShowDialog() = DialogResult.OK Then
                    Workbook.SaveAs(SFD.FileName)
                    MessageBox.Show("Exported File Saved to " & vbCrLf & SFD.FileName, "Save Exported File")
                End If

                Form1.ToolStripStatusLabel2.Text = "Complete"
                Form1.Cursor = Cursors.Default
            End Using

            ' Housekeeping
            Workbook.Close()
            xlApp.Quit()
            ReleaseObject(Worksheet)
            ReleaseObject(Workbook)
            ReleaseObject(xlApp)

            If Not Worksheet Is Nothing Then
                Marshal.FinalReleaseComObject(Worksheet)
                Worksheet = Nothing
            End If

            If Not Workbook Is Nothing Then
                Marshal.FinalReleaseComObject(Workbook)
                Workbook = Nothing
            End If

            If Not xlApp Is Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If

            'Last ditch to kill Excel
            For Each Proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
                Proc.Kill()
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
            xlApp.Quit()

        End Try

    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        '**********************************************
        '  Public Sub to Release the COM Object
        '**********************************************
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
        Catch ex As Exception
            MsgBox("Error releasing object" & ex.ToString)
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    Public Sub Export_Grid_Text_File(DGV As DataGridView)
        '*****************************************
        'Public Sub to Export the Grid view to a Text File
        '*****************************************
        Dim colIndex As Integer = 0

        Try
            'Save the Text file to a user location
            Using SFD As New SaveFileDialog
                If SFD.ShowDialog() = DialogResult.OK Then
                    DGV.AllowUserToAddRows = False
                    Using writer As New StreamWriter(SFD.FileName)

                        'Column Headers
                        For Each dc In DGV.Columns
                            If colIndex > 0 Then writer.Write(",")
                            colIndex = colIndex + 1
                            writer.Write(Chr(34) & "{0}" & Chr(34), dc.Name)
                        Next
                        writer.WriteLine()
                        'Write a line for each Row of the Grid
                        For iRow As Integer = 0 To DGV.Rows.Count - 1
                            For iCol As Integer = 0 To DGV.Columns.Count - 1
                                If iCol > 0 Then writer.Write(",")
                                writer.Write(Chr(34) & "{0}" & Chr(34), DGV.Rows(iRow).Cells(iCol).Value)
                            Next
                            writer.WriteLine()
                        Next
                        writer.Close()
                    End Using
                    MessageBox.Show("Exported Text File Saved to " & vbCrLf & SFD.FileName, "Save Exported File")
                    DGV.AllowUserToAddRows = True
                End If
            End Using

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

End Module
