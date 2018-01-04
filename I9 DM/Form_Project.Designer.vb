<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Project
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CancelProjectCmd = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ImageLocationText = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ProdComKey = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.VersionComBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProjectDescTx = New System.Windows.Forms.TextBox()
        Me.JIRATxt = New System.Windows.Forms.TextBox()
        Me.DemoComKeyTxt = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProjectLabel = New System.Windows.Forms.Label()
        Me.ProjectNameTxt = New System.Windows.Forms.TextBox()
        Me.AddNewProjectBut = New System.Windows.Forms.Button()
        Me.ProjectComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CancelProjectCmd
        '
        Me.CancelProjectCmd.BackColor = System.Drawing.Color.SkyBlue
        Me.CancelProjectCmd.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelProjectCmd.Location = New System.Drawing.Point(285, 357)
        Me.CancelProjectCmd.Name = "CancelProjectCmd"
        Me.CancelProjectCmd.Size = New System.Drawing.Size(75, 29)
        Me.CancelProjectCmd.TabIndex = 9
        Me.CancelProjectCmd.Text = "Cancel"
        Me.CancelProjectCmd.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ImageLocationText)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.ProdComKey)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.VersionComBox)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ProjectDescTx)
        Me.Panel1.Controls.Add(Me.JIRATxt)
        Me.Panel1.Controls.Add(Me.DemoComKeyTxt)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.ProjectLabel)
        Me.Panel1.Controls.Add(Me.ProjectNameTxt)
        Me.Panel1.Location = New System.Drawing.Point(12, 36)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(348, 303)
        Me.Panel1.TabIndex = 13
        '
        'ImageLocationText
        '
        Me.ImageLocationText.Location = New System.Drawing.Point(12, 194)
        Me.ImageLocationText.Name = "ImageLocationText"
        Me.ImageLocationText.Size = New System.Drawing.Size(319, 31)
        Me.ImageLocationText.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(9, 173)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(202, 28)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Image Location Path"
        '
        'ProdComKey
        '
        Me.ProdComKey.Location = New System.Drawing.Point(196, 89)
        Me.ProdComKey.Name = "ProdComKey"
        Me.ProdComKey.Size = New System.Drawing.Size(133, 31)
        Me.ProdComKey.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(193, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(186, 28)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Prod Company Key"
        '
        'VersionComBox
        '
        Me.VersionComBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.VersionComBox.FormattingEnabled = True
        Me.VersionComBox.Items.AddRange(New Object() {"G1", "G2"})
        Me.VersionComBox.Location = New System.Drawing.Point(217, 144)
        Me.VersionComBox.Name = "VersionComBox"
        Me.VersionComBox.Size = New System.Drawing.Size(86, 32)
        Me.VersionComBox.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(203, 123)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(172, 28)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Guardian Version"
        '
        'ProjectDescTx
        '
        Me.ProjectDescTx.Location = New System.Drawing.Point(10, 246)
        Me.ProjectDescTx.Multiline = True
        Me.ProjectDescTx.Name = "ProjectDescTx"
        Me.ProjectDescTx.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.ProjectDescTx.Size = New System.Drawing.Size(321, 45)
        Me.ProjectDescTx.TabIndex = 7
        '
        'JIRATxt
        '
        Me.JIRATxt.Location = New System.Drawing.Point(10, 144)
        Me.JIRATxt.Name = "JIRATxt"
        Me.JIRATxt.Size = New System.Drawing.Size(133, 31)
        Me.JIRATxt.TabIndex = 4
        '
        'DemoComKeyTxt
        '
        Me.DemoComKeyTxt.Location = New System.Drawing.Point(10, 89)
        Me.DemoComKeyTxt.Name = "DemoComKeyTxt"
        Me.DemoComKeyTxt.Size = New System.Drawing.Size(133, 31)
        Me.DemoComKeyTxt.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(7, 225)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(192, 28)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Description \ Notes"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(7, 123)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(203, 28)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "JIRA Project Number"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(7, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(197, 28)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Demo Company Key"
        '
        'ProjectLabel
        '
        Me.ProjectLabel.AutoSize = True
        Me.ProjectLabel.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProjectLabel.Location = New System.Drawing.Point(7, 12)
        Me.ProjectLabel.Name = "ProjectLabel"
        Me.ProjectLabel.Size = New System.Drawing.Size(205, 28)
        Me.ProjectLabel.TabIndex = 6
        Me.ProjectLabel.Text = "Create a new Project"
        '
        'ProjectNameTxt
        '
        Me.ProjectNameTxt.Location = New System.Drawing.Point(10, 33)
        Me.ProjectNameTxt.Name = "ProjectNameTxt"
        Me.ProjectNameTxt.Size = New System.Drawing.Size(321, 31)
        Me.ProjectNameTxt.TabIndex = 0
        '
        'AddNewProjectBut
        '
        Me.AddNewProjectBut.BackColor = System.Drawing.Color.SkyBlue
        Me.AddNewProjectBut.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddNewProjectBut.Location = New System.Drawing.Point(12, 357)
        Me.AddNewProjectBut.Name = "AddNewProjectBut"
        Me.AddNewProjectBut.Size = New System.Drawing.Size(84, 29)
        Me.AddNewProjectBut.TabIndex = 8
        Me.AddNewProjectBut.Text = "Create New"
        Me.AddNewProjectBut.UseVisualStyleBackColor = False
        '
        'ProjectComboBox
        '
        Me.ProjectComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ProjectComboBox.FormattingEnabled = True
        Me.ProjectComboBox.Location = New System.Drawing.Point(12, 7)
        Me.ProjectComboBox.Name = "ProjectComboBox"
        Me.ProjectComboBox.Size = New System.Drawing.Size(348, 32)
        Me.ProjectComboBox.TabIndex = 19
        Me.ProjectComboBox.Visible = False
        '
        'Form_Project
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(372, 400)
        Me.Controls.Add(Me.ProjectComboBox)
        Me.Controls.Add(Me.AddNewProjectBut)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CancelProjectCmd)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form_Project"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create a New Project"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CancelProjectCmd As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ProjectDescTx As TextBox
    Friend WithEvents JIRATxt As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents ProjectLabel As Label
    Friend WithEvents ProjectNameTxt As TextBox
    Friend WithEvents VersionComBox As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents AddNewProjectBut As Button
    Friend WithEvents ProjectComboBox As ComboBox
    Friend WithEvents ImageLocationText As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents ProdComKey As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents DemoComKeyTxt As TextBox
    Friend WithEvents Label3 As Label
End Class
