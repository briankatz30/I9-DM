
Module M_Variables
    '**********************************************
    ' Public Variables for the entire Project
    '**********************************************

    ' Project Variables
    Public RosterProjectTable, I9ProjectTable, GuardianVersion As String
    Public ProjectID As Integer
    Public RosterTableExsists, I9TableExsists, MakeTable As Boolean
    Public RosterMatch As Integer
    Public RosterNotMatchCount As Integer
    Public RosterExcelFilePath As String

    ' Database Variables
    'Public Variables for Main Project Db Connection
    Public Main_Conn As String = "Provider = SQLOLEDB.1; Data Source=PHX-035555; Trusted_Connection=True; Initial Catalog=LLX_PROJECT;User ID=LLX;Password=Password;"
    'For JumpBox Logon
    'Public Main_Conn As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=LLX_Project;Data Source=PHX-VMJUMP2\SQLEXPRESS;"
    Public Client_Conn As String

    'Image Variables
    Public Const ServerShare As String = "\\10.1.36.3\dmshare\Companies\"
    Public ImagePath As String
    Public DbLocation As String = IO.Path.GetFullPath(Application.StartupPath) & "\I9Project.mdb"
    Public BlankPDF As String = IO.Path.GetFullPath(Application.StartupPath) & "\Resources\Blank.pdf"

    'Variables for Image Array 
    Public ImageArray As New List(Of String)()
    Public CurrentElement As Integer

End Module
