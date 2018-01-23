'*************************************************************
' Module to Load data from text file into the Roster or I9 Tables
'*************************************************************
Option Explicit On
Module M_Load_Files

    Public Sub Load_Trans_File(ByVal FileName As String)
        '****************************************************************
        ' Sub that loads the user selected I9 file into the I9' table (Pipe Delimiter File)
        '****************************************************************
        Dim Rs As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String

        Dim LineofFile As String
        Dim FSO As New Scripting.FileSystemObject
        Dim Fs As Scripting.TextStream
        Dim FsLineCount As Scripting.TextStream
        Dim z, d, t As Integer
        Dim ReadLine() As String
        Dim TotalLines As Integer

        'Client Connection String
        Roster_Connection = Client_Conn

        Try
            'Need to open the Recordset to put the file into the temp table
            RosterConnection.Open(Roster_Connection)
            Rs.Open("SELECT * FROM [I9]  ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            'Reads the Lines of the file before we import
            FsLineCount = FSO.OpenTextFile(FileName, IOMode:=Scripting.IOMode.ForReading, Create:=False, Format:=Scripting.Tristate.TristateFalse)

            'Reads every line of the File to import
            Do Until FsLineCount.AtEndOfStream
                FsLineCount.ReadLine()
                TotalLines = TotalLines + 1
            Loop

            'Destroy the Object
            FsLineCount = Nothing

            'Progress Bar Display
            Form1.ToolStripStatusLabel2.Text = "Processing I9 Text File - " & TotalLines - 1 & " of " & TotalLines - 1

            'Opens the file again for import
            Fs = FSO.OpenTextFile(FileName, IOMode:=Scripting.IOMode.ForReading, Create:=False, Format:=Scripting.Tristate.TristateFalse)

            'Counters for Display
            d = 1

            Do Until Fs.AtEndOfStream
                'Split the Line of the Tab file
                Rs.AddNew()
                LineofFile = Fs.ReadLine
                LineofFile = Replace(LineofFile, Chr(34), "")
                'Reads a Tab Delimited File
                ReadLine = Split(LineofFile, "|")
                t = 1
                For z = 0 To UBound(ReadLine) - 1
                    Rs.Fields.Item(t).Value = ReadLine(z)
                    t = t + 1
                Next z
                Rs.Update()
                Form1.ToolStripStatusLabel2.Text = "Processing I9 Text File - " & TotalLines - 1 & " of " & d
                d = d + 1
            Loop
            Rs.Close()

            'HouseKeeping
            Rs = Nothing
            Fs = Nothing
            FSO = Nothing
            RosterConnection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Public Sub Load_Roster_File(ByVal FileName As String)
        '*******************************************************
        ' Sub Routine to load the user selected Roster file (Tab Delimited) 
        ' in to the Roster Table
        '*******************************************************
        Dim Rs As New ADODB.Recordset
        Dim RosterConnection As New ADODB.Connection
        Dim Roster_Connection As String
        Dim LineofFile As String

        'File Variables
        Dim FSO As New Scripting.FileSystemObject
        Dim Fs As Scripting.TextStream
        Dim FsLineCount As Scripting.TextStream
        Dim ReadLine() As String
        Dim TotalLines As Integer
        Dim z, d, t As Integer
        Dim LineNo As Integer = 1

        'Client Db Connection String
        Roster_Connection = Client_Conn

        Try
            'Need to open the Recordset to put the file into the temp table
            RosterConnection.Open(Roster_Connection)
            Rs.Open("SELECT * FROM [ROSTER] ;", RosterConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            'Reads the Lines of the file before we import
            FsLineCount = FSO.OpenTextFile(FileName, IOMode:=Scripting.IOMode.ForReading, Create:=False, Format:=Scripting.Tristate.TristateFalse)

            'Reads every line of the File to import for the display of importing records
            Do Until FsLineCount.AtEndOfStream
                FsLineCount.ReadLine()
                TotalLines = TotalLines + 1
            Loop

            'Destroy the Object
            FsLineCount = Nothing

            'Progress Bar Display
            Form1.ToolStripStatusLabel2.Text = "Processing Roster Text File - " & TotalLines - 1 & " of " & TotalLines - 1

            'Opens the file again for import
            Fs = FSO.OpenTextFile(FileName, IOMode:=Scripting.IOMode.ForReading, Create:=False, Format:=Scripting.Tristate.TristateFalse)

            'Counters for Display
            d = 1

            Do Until Fs.AtEndOfStream
                'Split the Line of the Tab file
                LineofFile = Fs.ReadLine
                LineofFile = Replace(LineofFile, Chr(34), "")
                t = 1
                'Reads the Tab Delimited File
                ReadLine = Split(LineofFile, Chr(9))
                'Remove the header Row
                If LineNo > 1 Then
                    Rs.AddNew()
                    For z = 0 To UBound(ReadLine)
                        Rs.Fields.Item(t).Value = ReadLine(z)
                        t = t + 1
                    Next z
                    Rs.Update()

                    d = d + 1
                    Form1.ToolStripStatusLabel2.Text = "Processing RosterText File - " & TotalLines - 1 & " of " & d
                Else
                    LineNo = LineNo + 1
                End If
            Loop

            'HouseKeeping
            Rs = Nothing
            Fs = Nothing
            FSO = Nothing
            RosterConnection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

End Module
