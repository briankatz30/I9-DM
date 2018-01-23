
Module M_LoadCombo
    Public Sub Load_Combo_Projects()
        '*******************************************************
        ' Sub to load the Projects into the Combo Boxes on the Main Form
        ' and the Combo Box on the Project Form reading the Main Project
        ' Database to load the Project Names into the ComboBoxes
        '*******************************************************
        Dim MyConn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset
        Dim SQLStr As String
        MyConn.Open(Main_Conn)
        SQLStr = "SELECT * FROM MAIN_LLX_PROJECT ORDER BY 2"
        Rs.Open(SQLStr, MyConn, CursorTypeEnum.adOpenKeyset)

        'Load the Project Combo Box on the tool strip with Project Names
        Form1.TSProjectComboBox.Items.Clear()
        Form_Project.ProjectComboBox.Items.Clear()
        'Loop to Load from the table MAIN_LLX_PROJECT
        Do While Not Rs.EOF
            Form1.TSProjectComboBox.Items.Add(Rs.Fields.Item("PROJECT_NAME").Value)
            Form_Project.ProjectComboBox.Items.Add(Rs.Fields.Item("PROJECT_NAME").Value)
            Rs.MoveNext()
        Loop

        'HouseKeeping
        Rs.Close()
        MyConn.Close()

    End Sub

    Public Sub Clean_Project_Form()
        '**********************************************
        ' Sub to remove any entries on the Form_Project Form
        ' before it loads
        '**********************************************
        Form_Project.ProjectNameTxt.Text = ""
        Form_Project.DemoComKeyTxt.Text = ""
        Form_Project.JIRATxt.Text = ""
        Form_Project.VersionComBox.Text = ""
        Form_Project.ProjectDescTx.Text = ""
        Form_Project.ProjectComboBox.Text = ""
        Form_Project.ProdComKey.Text = ""

    End Sub

End Module
