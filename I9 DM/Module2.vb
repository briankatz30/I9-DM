Module Module2
    Public Sub Delete_Temp_Table()
        'Delete the G2-Temp Table

        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        Dim SQLStmt As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation


        Connection = New OleDbConnection(ConnectionString)
        Connection.Open()
        SQLStmt = "DROP TABLE [G2-Temp]"
        Dim Cmd1 As OleDbCommand = New OleDbCommand(SQLStmt, Connection)
        Cmd1.ExecuteNonQuery()
        Cmd1.Dispose()

    End Sub
    ''Need to update the I-9 Folder with the location of the Images
    ''Opens the Folder Dialog window to select a path for the images
    ''Users Selects the folder path
    'If Form1.FolderBrowserDialogImages.ShowDialog() = DialogResult.OK Then
    '    ImageFolderPath = Form1.FolderBrowserDialogImages.SelectedPath
    '    'Need to truncate the Selected path to just from the Company Folder down
    '    '\\10.1.36.3\dmshare\Companies
    '    'Update the table with the new path
    '    Rs.Open("SELECT [I-9 FOLDER] FROM [" & QCTable & "] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    '    Rs.MoveFirst()
    '    ImageFolderPath = ""

    '    Do While Not Rs.EOF
    '        Rs.Fields.Item("I-9 FOLDER").Value = "\\10.1.36.3\dmshare\Companies" & ImageFolderPath & "\"
    '        Rs.Update()
    '        Rs.MoveNext()
    '    Loop
    'End If
    Public Sub Create_Temp_Roster_Table()
        '**********************************************
        '  Sub to Create either for G1 or G2
        '  Temp Table for the Roster File to be loaded into
        '**********************************************
        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim G1Temp, G2Temp As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()
            If GuardianVersion = "G1" Then
                'Create G1-Temp table
                G1Temp = "CREATE TABLE [G1-Temp]  " &
                " ([Employee ID] TEXT(255), [Employee Last Name] TEXT(255),  [Employee First Name] TEXT(255)," &
                " [Employee Middle Name] TEXT(255), [Employee Maiden Name] TEXT(255), [Employee Title] TEXT(255), " &
                " [Employee Date of Birth] TEXT(255), [Employee SS#] TEXT(255), [Employee Address] TEXT(255), " &
                " [Employee Address 2] TEXT(255), [Employee Apt #] TEXT(255),  [Employee City] TEXT(255),  [Employee State] TEXT(255), " &
                " [Employee Zip] TEXT(255), [Employee Country] TEXT(255), [Work Phone] TEXT(255), [Work Extension] TEXT(255), " &
                " [Home Phone] TEXT(255), [Home Extension] TEXT(255), [Cell Phone] TEXT(255), [Cell Extension] TEXT(255)," &
                " [Email Address] TEXT(255), [Location Name] TEXT(255), [Location Number] TEXT(255), [Occupation Class] TEXT(255), " &
                " [Business Unit] TEXT(255), [Hire Date] TEXT(255), [Terminated Date] Text(255)); "
                Dim Cmd As OleDbCommand = New OleDbCommand(G1Temp, Connection)
                Cmd.ExecuteNonQuery()
                Cmd.Dispose()
                Connection.Dispose()

            ElseIf GuardianVersion = "G2" Then
                'Create G2 Temp table
                G2Temp = "CREATE TABLE [G2-Temp]  " &
                " ([Employee ID] TEXT(255), [Employee Last Name] TEXT(255), [Employee First Name] TEXT(255),  [Employee Middle Name] TEXT(255), " &
                " [Employee Maiden Name] TEXT(255), [Employee Title] TEXT(255), [Employee Date of Birth] TEXT(255), [Employee SS#] TEXT(255), [Employee Address] TEXT(255), " &
                " [Employee Address 2] TEXT(255), [Employee Apt #] TEXT(255), [Employee City] TEXT(255), [Employee State] TEXT(255), [Employee Zip] TEXT(255)," &
                " [Employee Country] TEXT(255), [Work Phone] TEXT(255), [Work Extension] TEXT(255), [Home Phone] TEXT(255), [Home Extension] TEXT(255), " &
                " [Cell Phone] TEXT(255), [Cell Extension] TEXT(255), [Email Address] TEXT(255),  [Location Name] TEXT(255),  [Location Number] TEXT(255),  [Business Unit] TEXT(255), " &
                " [Hire Date] TEXT(255),  [Terminated Date] TEXT(255), [Custom Field 1] TEXT(255),  [Custom Field 2] TEXT(255),  [Custom Field 3] TEXT(255)," &
                " [Custom Field 4] TEXT(255),  [Custom Field 5] TEXT(255),  [Custom Field 6] TEXT(255),  [Custom Field 7] TEXT(255),  [Custom Field 8] TEXT(255));"
                Dim Cmd As OleDbCommand = New OleDbCommand(G2Temp, Connection)
                Cmd.ExecuteNonQuery()
                Cmd.Dispose()
                Connection.Close()
                Connection.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Load_G2_Data(ByVal G2Table As String, ByVal G2ClientTable As String)
        '**********************************************
        ' Sub to load data from the G2-Temp table to the
        ' Project Created by the user
        '**********************************************
        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim SQLStmt As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()
            'SQL Statement to move data from the G2 Temp Table to the new Roster table
            SQLStmt = "INSERT INTO [" & G2ClientTable & "] ( [Employee ID], [Employee Last Name], [Employee First Name], " &
            "[Employee Middle Name], [Employee Maiden Name], [Employee Title], [Employee Date of Birth], [Employee SS#], " &
            "[Employee Address], [Employee Address 2], [Employee Apt #], [Employee City], [Employee State], [Employee ZIP], " &
            "[Employee Country], [Work Phone], [Work Extension], [Home Phone], [Home Extension], [Cell Phone], [Cell Extension], " &
            "[Email Address], [Location Name], [Location Number], [Business Unit], [Hire Date], [Terminated Date],[Custom Field 1],[Custom Field 2], " &
            "[Custom Field 3], [Custom Field 4],[Custom Field 5],[Custom Field 6],[Custom Field 7],[Custom Field 8]) " &
            "SELECT [" & G2Table & "].[Employee ID], [" & G2Table & "].[Employee Last Name], [" & G2Table & "].[Employee First Name], [" & G2Table & "].[Employee Middle Name], " &
            "[" & G2Table & "].[Employee Maiden Name], [" & G2Table & "].[Employee Title], [" & G2Table & "].[Employee Date of Birth], [" & G2Table & "].[Employee SS#], " &
            "[" & G2Table & "].[Employee Address], [" & G2Table & "].[Employee Address 2], [" & G2Table & "].[Employee Apt #], [" & G2Table & "].[Employee City], " &
            "[" & G2Table & "].[Employee State], [" & G2Table & "].[Employee ZIP], [" & G2Table & "].[Employee Country], [" & G2Table & "].[Work Phone], " &
            "[" & G2Table & "].[Work Extension], [" & G2Table & "].[Home Phone], [" & G2Table & "].[Home Extension], [" & G2Table & "].[Cell Phone], [" & G2Table & "].[Cell Extension], " &
            "[" & G2Table & "].[Email Address], [" & G2Table & "].[Location Name], [" & G2Table & "].[Location Number], " &
            "[" & G2Table & "].[Business Unit], [" & G2Table & "].[Hire Date], [" & G2Table & "].[Terminated Date] , [" & G2Table & "].[Custom Field 1], " &
            "[" & G2Table & "].[Custom Field 2],[" & G2Table & "].[Custom Field 3],[" & G2Table & "].[Custom Field 4],[" & G2Table & "].[Custom Field 5]," &
            "[" & G2Table & "].[Custom Field 6],[" & G2Table & "].[Custom Field 7],[" & G2Table & "].[Custom Field 8] FROM " &
            "[" & G2Table & "] WHERE [EMPLOYEE ID] <> 'EMPLOYEE ID' ;"

            Dim Cmd As OleDbCommand = New OleDbCommand(SQLStmt, Connection)
            Cmd.ExecuteNonQuery()
            Cmd.Dispose()
            Connection.Close()
            Connection.Dispose()

            'Delete the G2-Temp Table
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()
            SQLStmt = "DROP TABLE [G2-Temp]"
            Dim Cmd1 As OleDbCommand = New OleDbCommand(SQLStmt, Connection)
            Cmd1.ExecuteNonQuery()
            Cmd1.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    '    Public Function Load_G1_Data(ByVal G1Table As String, ByVal G1ClientTable As String)
    '        '**********************************************
    '        ' Function to load data from the G1-Temp table to the
    '        ' Project Created by the user
    '        '**********************************************

    '        Dim SQLStmt As String

    '        'SQL Statement to move data from the G1 Temp Table to the new Client table
    '        SQLStmt = "INSERT INTO [" & G1ClientTable & "] ( [Employee ID], [Employee Last Name], [Employee First Name], " &
    '"[Employee Middle Name], [Employee Maiden Name], [Employee Title], [Employee Date of Birth], [Employee SS#], " &
    '"[Employee Address], [Employee Address 2], [Employee Apt #], [Employee City], [Employee State], [Employee ZIP], " &
    '"[Employee Country], [Work Phone], [Work Extension], [Home Phone], [Home Extension], [Cell Phone], [Cell Extension], " &
    '"[Email Address], [Location Name], [Location Number], [Occupation Class], [Business Unit], [Hire Date], [Terminated Date] ) " &
    '"SELECT [" & G1Table & "].[Employee ID], [" & G1Table & "].[Employee Last Name], [" & G1Table & "].[Employee First Name], [" & G1Table & "].[Employee Middle Name], " &
    '"[" & G1Table & "].[Employee Maiden Name], [" & G1Table & "].[Employee Title], [" & G1Table & "].[Employee Date of Birth], [" & G1Table & "].[Employee SS#], " &
    '"[" & G1Table & "].[Employee Address], [" & G1Table & "].[Employee Address 2], [" & G1Table & "].[Employee Apt #], [" & G1Table & "].[Employee City], " &
    '"[" & G1Table & "].[Employee State], [" & G1Table & "].[Employee ZIP], [" & G1Table & "].[Employee Country], [" & G1Table & "].[Work Phone], " &
    '"[" & G1Table & "].[Work Extension], [" & G1Table & "].[Home Phone], [" & G1Table & "].[Home Extension], [" & G1Table & "].[Cell Phone], [" & G1Table & "].[Cell Extension], " &
    '"[" & G1Table & "].[Email Address], [" & G1Table & "].[Location Name], [" & G1Table & "].[Location Number], [" & G1Table & "].[Occupation Class]," &
    '"[" & G1Table & "].[Business Unit], [" & G1Table & "].[Hire Date], [" & G1Table & "].[Terminated Date] FROM [" & G1Table & "];"

    '        db.Execute SQLStmt

    'End Function
    Public Sub Create_Temp_Trans_Table()
        '**********************************************
        '  Temp Table for the Transcription File to be loaded
        '**********************************************
        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim TransTemp As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()

            'Create Temp table
            TransTemp = "CREATE TABLE [TransTemp] ([2 page Flag] TEXT(255), [Section 3 Flag] TEXT(255),[Employee Last Name] TEXT(255),[Employee First Name] TEXT(255),[Employee Middle Initial] TEXT(255),[Employee Maiden Name] TEXT(255)," &
            " [Employee Address] TEXT(255),[Employee Apt #] TEXT(255),[Employee Date of Birth] TEXT(255),[Employee City] TEXT(255),[Employee State] TEXT(255),[Employee Zip] TEXT(255),[Employee SS#] TEXT(255),[Employee E-mail] TEXT(255), " &
            " [Employee Phone Number] TEXT(255),[Employee Status] TEXT(255),[Employee LPR Alien #] TEXT(255),[Employee Alien Authorized to Work Until] TEXT(255),[Employee Alien Registration/USCIS #] TEXT(255),[Employee Alien or Admission #] TEXT(255), " &
            " [Employee Foreign Passport #] TEXT(255), [Employee Foreign Passport Country] TEXT(255), [Employee Signature] TEXT(255),[Employee Signed Date] TEXT(255),[Translator Employee Setting] TEXT(255),[Translator Signature] TEXT(255)," &
            " [Translator Print Name/Last Name] TEXT(255),[Translator First Name] TEXT(255),[Translator Address] TEXT(255),[Translator City] TEXT(255),[Translator State] TEXT(255),[Translator Zip] TEXT(255),[Translator Signed Date] TEXT(255), " &
            " [Section 2 Header Last Name] TEXT(255),[Section 2 Header First Name] TEXT(255),[Section 2 Header Middle Initial] TEXT(255),[Section 2 Header Immigration Status] TEXT(255),[Document Title List A] TEXT(255),[Issuing Authority (A)] TEXT(255)," &
            " [Document Number (A1)] TEXT(255),[Expiration Date (A1)] TEXT(255),[Document Title List (A2)] TEXT(255),[Issuing Authority (A2)] TEXT(255),[Document Number (A2)] TEXT(255),[Expiration Date (A2)] TEXT(255),[Document List (A3)] TEXT(255), " &
            " [Issuing Authority (A3)] TEXT(255),[Document Number (A3)] TEXT(255),[Expiration Date (A3)] TEXT(255),[Document Title List (B)] TEXT(255),[Issuing Authority (B)] TEXT(255),[Document Number (B)] TEXT(255),[Expiration Date (B)] TEXT(255)," &
            " [Document Title List (C)] TEXT(255),[Issuing Authority (C)] TEXT(255),[Document Number (C)] TEXT(255),[Expiration Date (C)] TEXT(255),[Section 2 Addition Info] TEXT(255),[Employee Start Date] TEXT(255),[Supervisor Signature] TEXT(255), " &
            " [Supervisor Print/Last Name] TEXT(255),[Supervisor First Name] TEXT(255),[Supervisor Title] TEXT(255),[Business Name] TEXT(255),[Business Address] TEXT(255),[Business City] TEXT(255),[Business State] TEXT(255),[Business Zip] TEXT(255), " &
            " [Supervisor Signed Date] TEXT(255),[Employee New Last Name (Section 3)] TEXT(255),[Employee First Name (Section 3)] TEXT(255),[Employee Middle Initial (Section 3)] TEXT(255),[Date of Rehire (Section 3)] TEXT(255),[Document Title (Section 3)] TEXT(255), " &
            " [Document Number (Section 3)] TEXT(255),[Document Expiration Date (Section 3)] TEXT(255),[Supervisor Signature (Section 3)] TEXT(255),[Supervisor Signed Date (Section 3)] TEXT(255),[Supervisor Print Name (Section 3)] TEXT(255)," &
            " [Handwritten data in margins] TEXT(255),[Form Version] TEXT(255),[I-9 Folder] TEXT(255),[I-9 Document Name] TEXT(255),[I-9 Document Name 2] TEXT(255),[Supporting Doc 1 Name] TEXT(255),[Supporting Doc 2 Name] TEXT(255), " &
            " [Supporting Doc 3 Name] TEXT(255),[Supporting Doc 4 Name] TEXT(255),[Supporting Doc 5 Name] TEXT(255),[Supporting Doc 6 Name] TEXT(255),[Supporting Doc 7 Name] TEXT(255),[Supporting Doc 8 Name] TEXT(255),[Supporting Doc 9 Name] TEXT(255), " &
            " [Supporting Doc 10 Name] TEXT(255),[Supporting Doc 11 Name] TEXT(255),[Supporting Doc 12 Name] TEXT(255),[Supporting Doc 13 Name] TEXT(255),[Supporting Doc 14 Name] TEXT(255),[Supporting Doc 15 Name] TEXT(255),[Supporting Doc 16 Name] TEXT(255), " &
            " [Supporting Doc 17 Name] TEXT(255),[Supporting Doc 18 Name] TEXT(255),[Supporting Doc 19 Name] TEXT(255),[Supporting Doc 20 Name] TEXT(255),[Supporting Doc 21 Name] TEXT(255),[Supporting Doc 22 Name] TEXT(255))"

            Dim Cmd As OleDbCommand = New OleDbCommand(TransTemp, Connection)
            Cmd.ExecuteNonQuery()
            Cmd.Dispose()
            Connection.Close()
            Connection.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub Load_Trans_Data(ByVal TransTempTable As String, ByVal TransClientTable As String)
        '**********************************************
        ' Sub to load data from the TransTemp table to the
        ' Project Created by the user
        '**********************************************
        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim SQLStmt, SQLStmt1, SQLStmt2 As String
        Dim SQLStmts As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()

            'SQL Statement to move data from the G1 Temp Table to the new Client table
            SQLStmt = "INSERT INTO [" & TransClientTable & "] ([2 page Flag], [Section 3 Flag],[Employee Last Name],[Employee First Name],[Employee Middle Initial],[Employee Maiden Name],[Employee Address], " &
            " [Employee Apt #],[Employee Date of Birth],[Employee City],[Employee State],[Employee Zip],[Employee SS#],[Employee E-mail],[Employee Phone Number],[Employee Status],[Employee LPR Alien #], " &
            " [Employee Alien Authorized to Work Until],[Employee Alien Registration/USCIS #],[Employee Alien or Admission #],[Employee Foreign Passport #],[Employee Foreign Passport Country],[Employee Signature], " &
            " [Employee Signed Date],[Translator Employee Setting],[Translator Signature],[Translator Print Name/Last Name],[Translator First Name],[Translator Address],[Translator City],[Translator State],[Translator Zip], " &
            " [Translator Signed Date],[Section 2 Header Last Name],[Section 2 Header First Name],[Section 2 Header Middle Initial],[Section 2 Header Immigration Status],[Document Title List A],[Issuing Authority (A)], " &
            " [Document Number (A1)],[Expiration Date (A1)],[Document Title List (A2)],[Issuing Authority (A2)],[Document Number (A2)],[Expiration Date (A2)],[Document List (A3)],[Issuing Authority (A3)], " &
            " [Document Number (A3)],[Expiration Date (A3)],[Document Title List (B)],[Issuing Authority (B)],[Document Number (B)],[Expiration Date (B)],[Document Title List (C)],[Issuing Authority (C)], " &
            " [Document Number (C)],[Expiration Date (C)],[Section 2 Addition Info],[Employee Start Date],[Supervisor Signature],[Supervisor Print/Last Name],[Supervisor First Name],[Supervisor Title], " &
            " [Business Name],[Business Address],[Business City],[Business State],[Business Zip],[Supervisor Signed Date],[Employee New Last Name (Section 3)],[Employee First Name (Section 3)], " &
            " [Employee Middle Initial (Section 3)],[Date of Rehire (Section 3)],[Document Title (Section 3)],[Document Number (Section 3)],[Document Expiration Date (Section 3)],[Supervisor Signature (Section 3)], " &
            " [Supervisor Signed Date (Section 3)],[Supervisor Print Name (Section 3)],[Handwritten data in margins],[Form Version],[I-9 Folder],[I-9 Document Name],[I-9 Document Name 2],[Supporting Doc 1 Name], " &
            " [Supporting Doc 2 Name],[Supporting Doc 3 Name],[Supporting Doc 4 Name],[Supporting Doc 5 Name],[Supporting Doc 6 Name],[Supporting Doc 7 Name],[Supporting Doc 8 Name],[Supporting Doc 9 Name],[Supporting Doc 10 Name], " &
            " [Supporting Doc 11 Name],[Supporting Doc 12 Name],[Supporting Doc 13 Name],[Supporting Doc 14 Name],[Supporting Doc 15 Name],[Supporting Doc 16 Name],[Supporting Doc 17 Name], " &
            " [Supporting Doc 18 Name],[Supporting Doc 19 Name],[Supporting Doc 20 Name],[Supporting Doc 21 Name],[Supporting Doc 22 Name]) "

            SQLStmt1 = "SELECT  [" & TransTempTable & "].[2 page Flag],[" & TransTempTable & "].[Section 3 Flag],[" & TransTempTable & "].[Employee Last Name],[" & TransTempTable & "].[Employee First Name], " &
            " [" & TransTempTable & "].[Employee Middle Initial],[" & TransTempTable & "].[Employee Maiden Name],[" & TransTempTable & "].[Employee Address],[" & TransTempTable & "].[Employee Apt #],[" & TransTempTable & "].[Employee Date of Birth], " &
            " [" & TransTempTable & "].[Employee City],[" & TransTempTable & "].[Employee State],[" & TransTempTable & "].[Employee Zip],[" & TransTempTable & "].[Employee SS#],[" & TransTempTable & "].[Employee E-mail], " &
            " [" & TransTempTable & "].[Employee Phone Number],[" & TransTempTable & "].[Employee Status],[" & TransTempTable & "].[Employee LPR Alien #],[" & TransTempTable & "].[Employee Alien Authorized to Work Until],[" & TransTempTable & "].[Employee Alien Registration/USCIS #]," &
            " [" & TransTempTable & "].[Employee Alien or Admission #],[" & TransTempTable & "].[Employee Foreign Passport #],[" & TransTempTable & "].[Employee Foreign Passport Country],[" & TransTempTable & "].[Employee Signature], " &
            " [" & TransTempTable & "].[Employee Signed Date],[" & TransTempTable & "].[Translator Employee Setting],[" & TransTempTable & "].[Translator Signature],[" & TransTempTable & "].[Translator Print Name/Last Name],[" & TransTempTable & "].[Translator First Name], " &
            " [" & TransTempTable & "].[Translator Address],[" & TransTempTable & "].[Translator City],[" & TransTempTable & "].[Translator State],[" & TransTempTable & "].[Translator Zip],[" & TransTempTable & "].[Translator Signed Date], " &
            " [" & TransTempTable & "].[Section 2 Header Last Name],[" & TransTempTable & "].[Section 2 Header First Name],[" & TransTempTable & "].[Section 2 Header Middle Initial],[" & TransTempTable & "].[Section 2 Header Immigration Status], " &
            " [" & TransTempTable & "].[Document Title List A],[" & TransTempTable & "].[Issuing Authority (A)],[" & TransTempTable & "].[Document Number (A1)],[" & TransTempTable & "].[Expiration Date (A1)], " &
            " [" & TransTempTable & "].[Document Title List (A2)],[" & TransTempTable & "].[Issuing Authority (A2)],[" & TransTempTable & "].[Document Number (A2)],[" & TransTempTable & "].[Expiration Date (A2)],[" & TransTempTable & "].[Document List (A3)], " &
            " [" & TransTempTable & "].[Issuing Authority (A3)],[" & TransTempTable & "].[Document Number (A3)],[" & TransTempTable & "].[Expiration Date (A3)],[" & TransTempTable & "].[Document Title List (B)], " &
            " [" & TransTempTable & "].[Issuing Authority (B)],[" & TransTempTable & "].[Document Number (B)],[" & TransTempTable & "].[Expiration Date (B)],[" & TransTempTable & "].[Document Title List (C)],[" & TransTempTable & "].[Issuing Authority (C)],  " &
            " [" & TransTempTable & "].[Document Number (C)],[" & TransTempTable & "].[Expiration Date (C)],[" & TransTempTable & "].[Section 2 Addition Info],[" & TransTempTable & "].[Employee Start Date],[" & TransTempTable & "].[Supervisor Signature], " &
            " [" & TransTempTable & "].[Supervisor Print/Last Name],[" & TransTempTable & "].[Supervisor First Name],[" & TransTempTable & "].[Supervisor Title],[" & TransTempTable & "].[Business Name],[" & TransTempTable & "].[Business Address], " &
            " [" & TransTempTable & "].[Business City],[" & TransTempTable & "].[Business State],[" & TransTempTable & "].[Business Zip],[" & TransTempTable & "].[Supervisor Signed Date],[" & TransTempTable & "].[Employee New Last Name (Section 3)], " &
            " [" & TransTempTable & "].[Employee First Name (Section 3)],[" & TransTempTable & "].[Employee Middle Initial (Section 3)],[" & TransTempTable & "].[Date of Rehire (Section 3)],[" & TransTempTable & "].[Document Title (Section 3)], " &
            " [" & TransTempTable & "].[Document Number (Section 3)],[" & TransTempTable & "].[Document Expiration Date (Section 3)],[" & TransTempTable & "].[Supervisor Signature (Section 3)],[" & TransTempTable & "].[Supervisor Signed Date (Section 3)], " &
            " [" & TransTempTable & "].[Supervisor Print Name (Section 3)],[" & TransTempTable & "].[Handwritten data in margins],[" & TransTempTable & "].[Form Version],[" & TransTempTable & "].[I-9 Folder]," &
            " [" & TransTempTable & "].[I-9 Document Name],[" & TransTempTable & "].[I-9 Document Name 2],[" & TransTempTable & "].[Supporting Doc 1 Name],[" & TransTempTable & "].[Supporting Doc 2 Name],[" & TransTempTable & "].[Supporting Doc 3 Name], " &
            " [" & TransTempTable & "].[Supporting Doc 4 Name],[" & TransTempTable & "].[Supporting Doc 5 Name],[" & TransTempTable & "].[Supporting Doc 6 Name],[" & TransTempTable & "].[Supporting Doc 7 Name],[" & TransTempTable & "].[Supporting Doc 8 Name], " &
            " [" & TransTempTable & "].[Supporting Doc 9 Name],[" & TransTempTable & "].[Supporting Doc 10 Name],[" & TransTempTable & "].[Supporting Doc 11 Name],[" & TransTempTable & "].[Supporting Doc 12 Name],[" & TransTempTable & "].[Supporting Doc 13 Name], " &
            " [" & TransTempTable & "].[Supporting Doc 14 Name],[" & TransTempTable & "].[Supporting Doc 15 Name],[" & TransTempTable & "].[Supporting Doc 16 Name],[" & TransTempTable & "].[Supporting Doc 17 Name]," &
            " [" & TransTempTable & "].[Supporting Doc 18 Name],[" & TransTempTable & "].[Supporting Doc 19 Name],[" & TransTempTable & "].[Supporting Doc 20 Name],[" & TransTempTable & "].[Supporting Doc 21 Name], " &
            " [" & TransTempTable & "].[Supporting Doc 22 Name] "

            SQLStmt2 = "FROM [" & TransTempTable & "] ;"

            SQLStmts = SQLStmt & SQLStmt1 & SQLStmt2
            Dim Cmd As OleDbCommand = New OleDbCommand(SQLStmts, Connection)
            Cmd.ExecuteNonQuery()
            Cmd.Dispose()
            Connection.Close()
            Connection.Dispose()

            'Delete the G2-Temp Table
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()
            SQLStmt = "DROP TABLE [TransTemp]"
            Dim Cmd1 As OleDbCommand = New OleDbCommand(SQLStmt, Connection)
            Cmd1.ExecuteNonQuery()
            Cmd1.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Create_I9_Template(ByVal I9Project As String)
        '**********************************************
        ' Sub to Create User Table for G2 using the Template
        ' Project Created by the user
        '**********************************************

        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim SqlStmt As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()

            SqlStmt = "CREATE TABLE [" & I9Project & "] ([ID] COUNTER, [2 page Flag] TEXT(255), [Section 3 Flag] TEXT(255), [Employee Last Name] TEXT(255), [Employee First Name] TEXT(255)," &
            " [Employee Middle Initial] TEXT(255), [Employee Maiden Name] TEXT(255), [Employee Address] TEXT(255), [Employee Apt #] TEXT(255), [Employee Date of Birth] TEXT(255)," &
            " [Employee City] TEXT(255), [Employee State] TEXT(255), [Employee Zip] TEXT(255), [Employee SS#] TEXT(255), [Employee E-mail] TEXT(255), [Employee Phone Number] TEXT(255), " &
            " [Employee Status] TEXT(255), [Employee LPR Alien #] TEXT(255), [Employee Alien Authorized to Work Until] TEXT(255), [Employee Alien Registration/USCIS #] TEXT(255), " &
            " [Employee Alien Or Admission #] TEXT(255), [Employee Foreign Passport #] TEXT(255), [Employee Foreign Passport Country] TEXT(255), [Employee Signature] TEXT(255), " &
            " [Employee Signed Date] TEXT(255), [Translator Employee Setting] TEXT(255), [Translator Signature] TEXT(255), [Translator Print Name/Last Name] TEXT(255), " &
            " [Translator First Name] TEXT(255), [Translator Address] TEXT(255), [Translator City] TEXT(255), [Translator State] TEXT(255), [Translator Zip] TEXT(255)," &
            " [Translator Signed Date] TEXT(255), [Section 2 Header Last Name] TEXT(255),  [Section 2 Header First Name] TEXT(255), [Section 2 Header Middle Initial] TEXT(255), " &
            " [Section 2 Header Immigration Status] TEXT(255), [Document Title List A] TEXT(255), [Issuing Authority (A)] TEXT(255), [Document Number (A1)] TEXT(255)," &
            " [Expiration Date (A1)] TEXT(255), [Document Title List (A2)] TEXT(255), [Issuing Authority (A2)] TEXT(255), [Document Number (A2)] TEXT(255), " &
            " [Expiration Date (A2)] TEXT(255), [Document List (A3)] TEXT(255), [Issuing Authority (A3)] TEXT(255), [Document Number (A3)] TEXT(255), [Expiration Date (A3)] TEXT(255), " &
            " [Document Title List (B)] TEXT(255), [Issuing Authority (B)] TEXT(255), [Document Number (B)] TEXT(255), [Expiration Date (B)] TEXT(255), [Document Title List (C)] TEXT(255), " &
            " [Issuing Authority (C)] TEXT(255), [Document Number (C)] TEXT(255), [Expiration Date (C)] TEXT(255), [Section 2 Addition Info] TEXT(255), " &
            " [Employee Start Date] TEXT(255), [Supervisor Signature] TEXT(255), [Supervisor Print/Last Name] TEXT(255), [Supervisor First Name] TEXT(255), " &
            " [Supervisor Title] TEXT(255), [Business Name] TEXT(255), [Business Address] TEXT(255), [Business City] TEXT(255),  [Business State] TEXT(255), [Business Zip] TEXT(255), " &
            " [Supervisor Signed Date] TEXT(255), [Employee New Last Name (Section 3)] TEXT(255), [Employee First Name (Section 3)] TEXT(255), [Employee Middle Initial (Section 3)] TEXT(255)," &
            " [Date of Rehire (Section 3)] TEXT(255), [Document Title (Section 3)] TEXT(255), [Document Number (Section 3)] TEXT(255), [Document Expiration Date (Section 3)] TEXT(255), " &
            " [Supervisor Signature (Section 3)] TEXT(255), [Supervisor Signed Date (Section 3)] TEXT(255), [Supervisor Print Name (Section 3)] TEXT(255), [Handwritten data in margins] MEMO, " &
            " [Form Version] TEXT(255), [I-9 Folder] TEXT(255), [I-9 Document Name] MEMO, [I-9 Document Name 2] MEMO, [Supporting Doc 1 Name] MEMO, [Supporting Doc 2 Name] MEMO, " &
            " [Supporting Doc 3 Name] MEMO, [Supporting Doc 4 Name] MEMO, [Supporting Doc 5 Name] MEMO, [Supporting Doc 6 Name] MEMO, [Supporting Doc 7 Name] MEMO, " &
            " [Supporting Doc 8 Name] MEMO, [Supporting Doc 9 Name] MEMO, [Supporting Doc 10 Name] MEMO, [Supporting Doc 11 Name] MEMO, [Supporting Doc 12 Name] MEMO, " &
            " [Supporting Doc 13 Name] MEMO, [Supporting Doc 14 Name] MEMO, [Supporting Doc 15 Name] MEMO, [Supporting Doc 16 Name] MEMO, [Supporting Doc 17 Name] MEMO, " &
            " [Supporting Doc 18 Name] MEMO, [Supporting Doc 19 Name] MEMO, [Supporting Doc 20 Name] MEMO, [Supporting Doc 21 Name] MEMO, [Supporting Doc 22 Name] MEMO, " &
            " [MATCH] TEXT(20), [ROSTERID] SMALLINT)"

            Dim Cmd As OleDbCommand = New OleDbCommand(SqlStmt, Connection)
            Cmd.ExecuteNonQuery()
            Cmd.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Create_G2_Template(ByVal RosterProject As String)
        '**********************************************
        ' Sub to Create User Table for G2 using the Template
        ' Project Created by the user
        '**********************************************

        Dim Connection As OleDbConnection
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source= " & DbLocation
        Dim SqlStmt As String

        Try
            Connection = New OleDbConnection(ConnectionString)
            Connection.Open()

            SqlStmt = "CREATE TABLE [" & RosterProject & "] ([ID] COUNTER, [Employee ID] TEXT(255), [Employee Last Name] TEXT(255),  [Employee First Name] TEXT(255), " &
           "[Employee Middle Name] TEXT(255), [Employee Maiden Name] TEXT(255), [Employee Title] TEXT(255), [Employee Date of Birth] TEXT(255), [Employee SS#] TEXT(255), " &
           "[Employee Address] TEXT(255),  [Employee Address 2] TEXT(255),  [Employee Apt #] TEXT(255),  [Employee City] TEXT(255),  [Employee State] TEXT(255), " &
           "[Employee Zip] TEXT(255), [Employee Country] TEXT(255),  [Work Phone] TEXT(255), [Work Extension] TEXT(255), [Home Phone] TEXT(255), [Home Extension] TEXT(255), " &
           "[Cell Phone] TEXT(255), [Cell Extension] TEXT(255), [Email Address] TEXT(255), [Location Name] TEXT(255),  [Location Number] TEXT(255),  [Business Unit] TEXT(255), " &
           "[Hire Date] TEXT(255), [Terminated Date] TEXT(255), [Custom Field 1] TEXT(255), [Custom Field 2] TEXT(255), [Custom Field 3] TEXT(255), [Custom Field 4] TEXT(255), " &
           "[Custom Field 5] TEXT(255),  [Custom Field 6] TEXT(255),  [Custom Field 7] TEXT(255), [Custom Field 8] TEXT(255), [Date Error] TEXT(15), [Date Description] MEMO, " &
           "[SSN Error] TEXT(15), [SSN Description] MEMO, [Other Error] TEXT(15), [Other Description] MEMO)"

            Dim Cmd As OleDbCommand = New OleDbCommand(SqlStmt, Connection)
            Cmd.ExecuteNonQuery()
            Cmd.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Read_CSV_File()
        '    Dim SR As StreamReader = New StreamReader("your csv file path")
        '    Dim line As String = SR.ReadLine()
        '    Dim strArray As String() = line.Split(","c)
        '    Dim dt As DataTable = New DataTable()
        '    Dim row As DataRow

        '    For Each s As String In strArray
        '        dt.Columns.Add(New DataColumn())
        '    Next

        '    Do
        '        line = SR.ReadLine
        '        If Not line = String.Empty Then
        '            row = dt.NewRow()
        '            row.ItemArray = line.Split(","c)
        '            dt.Rows.Add(row)
        '        Else
        '            Exit Do
        '        End If
        '    Loop
        'End Sub

        'Using cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConsoleApplication3.Properties.Settings.daasConnectionString").ConnectionString)
        '    cn.Open()
        '    Using copy As New SqlBulkCopy(cn)
        '        copy.ColumnMappings.Add(0, 0)
        '        copy.ColumnMappings.Add(1, 1)
        '        copy.ColumnMappings.Add(2, 2)
        '        copy.ColumnMappings.Add(3, 3)
        '        copy.ColumnMappings.Add(4, 4)
        '        copy.DestinationTableName = "Censis"
        '        copy.WriteToServer(dt)
        '    End Using
        'End Using

    End Sub


End Module
