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


        'Need to loop though the Grid grabbing the Included checkbox and create an 
        'Update Statement based on the 










    End Sub


End Class