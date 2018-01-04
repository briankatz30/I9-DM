Module M_ViewImage

    Public Sub View_Image(ByVal RosterID As String, ByVal GridView As String)
        '*********************************************************************
        ' Sub Routine to get the image path of all the files and place them into an array for
        ' viewing them in the PDF Viewer.  Grabs which Grid the users is selecting  to know if the Grid is from the Roster or 
        ' Transcript
        ' ********************************************************************

        Dim Rs As New ADODB.Recordset
        Dim sqlStmt As String
        Dim OneImage As String
        'Dim Doc1 As String
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String

        Roster_Connection = Client_Conn

        Try
            'Need to open the Recordset to put the files into a recordset based on the Grid that is being selected
            RosterConnection.Open(Roster_Connection)
            Select Case GridView
                Case = "Roster"
                    'Roster Grid View
                    sqlStmt = "SELECT [I-9 Document Name], [I-9 Document Name 2], [Supporting Doc 1 Name], [Supporting Doc 2 Name], " &
                        "[Supporting Doc 3 Name], [Supporting Doc 4 Name], [Supporting Doc 5 Name],[Supporting Doc 6 Name],[Supporting Doc 7 Name],[Supporting Doc 8 Name], " &
                        "[Supporting Doc 9 Name], [Supporting Doc 10 Name], [Supporting Doc 11 Name],[Supporting Doc 12 Name],[Supporting Doc 13 Name],[Supporting Doc 14 Name], " &
                        "[Supporting Doc 15 Name], [Supporting Doc 16 Name],[Supporting Doc 17 Name],[Supporting Doc 18 Name],[Supporting Doc 19 Name],[Supporting Doc 20 Name], " &
                        "[Supporting Doc 21 Name], [Supporting Doc 22 Name], [I-9 FOLDER] FROM  [I9] WHERE [RosterID] = " & RosterID & " ORDER BY ID"
                    Rs.Open(sqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                Case = "Transaction"
                    'Transcription Grid View
                    sqlStmt = "SELECT [I-9 Document Name], [I-9 Document Name 2],[Supporting Doc 1 Name], [Supporting Doc 2 Name], " &
                    "[Supporting Doc 3 Name], [Supporting Doc 4 Name], [Supporting Doc 5 Name],[Supporting Doc 6 Name],[Supporting Doc 7 Name],[Supporting Doc 8 Name], " &
                    "[Supporting Doc 9 Name], [Supporting Doc 10 Name], [Supporting Doc 11 Name],[Supporting Doc 12 Name],[Supporting Doc 13 Name],[Supporting Doc 14 Name], " &
                    "[Supporting Doc 15 Name], [Supporting Doc 16 Name],[Supporting Doc 17 Name],[Supporting Doc 18 Name],[Supporting Doc 19 Name],[Supporting Doc 20 Name], " &
                    "[Supporting Doc 21 Name], [Supporting Doc 22 Name],  [I-9 FOLDER]  FROM  [I9] WHERE [ID] = " & RosterID & " ORDER BY ID"
                    Rs.Open(sqlStmt, RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            End Select

            Dim v As Integer
            ImageArray.Clear()
            Dim i As Integer = 0
            'Loads all the PDF Files into a an array
            Do While Not Rs.EOF
                v = Rs.RecordCount
                'Loop through each image field in the recordset
                For i = 0 To 23
                    If Not IsDBNull(Rs(i).Value) Then
                        'Field has a value
                        ImageArray.Add(ImagePath & Rs(i).Value)
                    Else
                        'Nothing in the field
                    End If
                Next i
                Rs.MoveNext()
            Loop

            'Now work on the array
            If ImageArray.Count > 0 Then
                'Means we have images in the record
                'Just one image to load
                If ImageArray.Count = 1 Then
                    'Just one image and we load that one to the screen
                    OneImage = ImageArray(0).ToString
                    Form1.TextBox114.Text = OneImage
                    Form1.AxAcroPDF1.LoadFile(OneImage)
                    'Form1.ImageI9.Enabled = False
                    Form1.ImagePage2.Enabled = False
                    Form1.ImageSupNext.Enabled = False
                    Form1.ImageSupPrev.Enabled = False
                    CurrentElement = ImageArray.IndexOf(OneImage)
                    Exit Sub
                End If
                If ImageArray.Count > 1 Then
                    'More than Two Pages
                    OneImage = ImageArray(0).ToString
                    Form1.TextBox114.Text = ImageArray(0).ToString
                    Form1.AxAcroPDF1.LoadFile(OneImage)
                    CurrentElement = ImageArray.IndexOf(OneImage)
                    Form1.ImagePage2.Enabled = True
                    Form1.ImageSupNext.Enabled = True
                    Form1.ImageSupPrev.Enabled = True
                    Exit Sub
                End If
            Else
                'Nothing in the record
                Form1.AxAcroPDF1.LoadFile(BlankPDF)
                Form1.TextBox114.Text = String.Empty
                'Form1.ImageI9.Enabled = False
                'Form1.ImagePage2.Enabled = False
                'Form1.ImageSupNext.Enabled = False
                'Form1.ImageSupPrev.Enabled = False
                Exit Sub
            End If

            'Console.WriteLine(ImageArray.Count)
            Rs.MoveFirst()

            If Rs.EOF Or Rs.BOF Then
                'Form1.AxAcroPDF1.LoadFile(BlankPDF)
                Exit Sub
                'No records that match
            Else
                'Check for Null is the fields so that we have a good file name to use
                Rs.MoveFirst()
                'Loads the document path
                'DocPath = Rs.Fields.Item("I-9 FOLDER").Value & Rs.Fields.Item("I-9 Document Name").Value
                'Form1.TextBox114.Text = DocPath.ToString
                'Form1.AxAcroPDF1.LoadFile(DocPath)
                '    If (IsDBNull(Rs.Fields.Item("I-9 Document Name"))) Then
                '        DocPath = BlankPDF
                '    Else
                '        DocPath = Rs.Fields.Item("I-9 FOLDER").Value & Rs.Fields.Item("I-9 Document Name").Value
                '    End If
                '    If (IsDBNull(Rs.Fields.Item("I-9 Document Name 2"))) Then
                '        'DocPath = "Blank.pdf"
                '    Else
                '        'Doc2 = Rs.Fields.Item("I-9 Document Name 2").Value
                '    End If
            End If

            Rs.Close()
            RosterConnection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Module
