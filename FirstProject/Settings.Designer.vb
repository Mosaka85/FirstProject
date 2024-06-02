Imports System.ComponentModel
Imports System.Drawing.Drawing2D
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Settings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Settings))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUsernameSQL = New System.Windows.Forms.TextBox()
        Me.lblusernameSQL = New System.Windows.Forms.Label()
        Me.txtdatabaseName = New System.Windows.Forms.TextBox()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSavesettings = New System.Windows.Forms.Button()
        Me.btnNewConnectionSettings = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnUsersTable = New System.Windows.Forms.Button()
        Me.txtRefreshtimer = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckRefeshTimer = New System.Windows.Forms.CheckBox()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ComboBox1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.ForeColor = System.Drawing.Color.Black
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(447, 299)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(278, 24)
        Me.ComboBox1.TabIndex = 0
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPassword.Font = New System.Drawing.Font("Arial", 13.8!)
        Me.txtPassword.ForeColor = System.Drawing.Color.Black
        Me.txtPassword.Location = New System.Drawing.Point(410, 197)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(8226)
        Me.txtPassword.Size = New System.Drawing.Size(507, 27)
        Me.txtPassword.TabIndex = 1
        Me.txtPassword.Text = "76yy"
        '
        'txtUsernameSQL
        '
        Me.txtUsernameSQL.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtUsernameSQL.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtUsernameSQL.Font = New System.Drawing.Font("Arial", 13.8!)
        Me.txtUsernameSQL.Location = New System.Drawing.Point(410, 129)
        Me.txtUsernameSQL.Name = "txtUsernameSQL"
        Me.txtUsernameSQL.Size = New System.Drawing.Size(507, 27)
        Me.txtUsernameSQL.TabIndex = 2
        '
        'lblusernameSQL
        '
        Me.lblusernameSQL.AutoSize = True
        Me.lblusernameSQL.BackColor = System.Drawing.Color.Transparent
        Me.lblusernameSQL.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblusernameSQL.Location = New System.Drawing.Point(277, 127)
        Me.lblusernameSQL.Name = "lblusernameSQL"
        Me.lblusernameSQL.Size = New System.Drawing.Size(127, 29)
        Me.lblusernameSQL.TabIndex = 3
        Me.lblusernameSQL.Text = "Username"
        '
        'txtdatabaseName
        '
        Me.txtdatabaseName.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtdatabaseName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtdatabaseName.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdatabaseName.Location = New System.Drawing.Point(410, 76)
        Me.txtdatabaseName.Name = "txtdatabaseName"
        Me.txtdatabaseName.Size = New System.Drawing.Size(507, 27)
        Me.txtdatabaseName.TabIndex = 4
        '
        'btnConnect
        '
        Me.btnConnect.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnConnect.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnConnect.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConnect.ForeColor = System.Drawing.Color.Black
        Me.btnConnect.Location = New System.Drawing.Point(842, 228)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(75, 23)
        Me.btnConnect.TabIndex = 5
        Me.btnConnect.Text = "Test"
        Me.btnConnect.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(277, 198)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 29)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(246, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(157, 29)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Server Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(246, 293)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(195, 29)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Select Database"
        '
        'btnSavesettings
        '
        Me.btnSavesettings.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSavesettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSavesettings.ForeColor = System.Drawing.Color.Black
        Me.btnSavesettings.Location = New System.Drawing.Point(842, 379)
        Me.btnSavesettings.Name = "btnSavesettings"
        Me.btnSavesettings.Size = New System.Drawing.Size(75, 33)
        Me.btnSavesettings.TabIndex = 9
        Me.btnSavesettings.Text = "Save"
        Me.btnSavesettings.UseVisualStyleBackColor = False
        '
        'btnNewConnectionSettings
        '
        Me.btnNewConnectionSettings.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnNewConnectionSettings.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNewConnectionSettings.ForeColor = System.Drawing.Color.Black
        Me.btnNewConnectionSettings.Location = New System.Drawing.Point(948, 12)
        Me.btnNewConnectionSettings.Name = "btnNewConnectionSettings"
        Me.btnNewConnectionSettings.Size = New System.Drawing.Size(113, 23)
        Me.btnNewConnectionSettings.TabIndex = 10
        Me.btnNewConnectionSettings.Text = "Clear Fields"
        Me.btnNewConnectionSettings.UseVisualStyleBackColor = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 538)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.ManagerRenderMode
        Me.StatusStrip1.Size = New System.Drawing.Size(1073, 26)
        Me.StatusStrip1.TabIndex = 11
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(316, 20)
        Me.ToolStripStatusLabel1.Text = "© 2024 MOSAKA LIMITED All Rights Reserved."
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(1969, 20)
        Me.ToolStripStatusLabel2.Text = resources.GetString("ToolStripStatusLabel2.Text")
        '
        'btnUsersTable
        '
        Me.btnUsersTable.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUsersTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUsersTable.ForeColor = System.Drawing.Color.Black
        Me.btnUsersTable.Location = New System.Drawing.Point(740, 290)
        Me.btnUsersTable.Name = "btnUsersTable"
        Me.btnUsersTable.Size = New System.Drawing.Size(80, 33)
        Me.btnUsersTable.TabIndex = 12
        Me.btnUsersTable.Text = "Setup DB"
        Me.btnUsersTable.UseVisualStyleBackColor = False
        '
        'txtRefreshtimer
        '
        Me.txtRefreshtimer.Location = New System.Drawing.Point(360, 355)
        Me.txtRefreshtimer.Name = "txtRefreshtimer"
        Me.txtRefreshtimer.Size = New System.Drawing.Size(100, 22)
        Me.txtRefreshtimer.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(257, 361)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "RefreshTime"
        '
        'CheckRefeshTimer
        '
        Me.CheckRefeshTimer.AutoSize = True
        Me.CheckRefeshTimer.Location = New System.Drawing.Point(467, 361)
        Me.CheckRefeshTimer.Name = "CheckRefeshTimer"
        Me.CheckRefeshTimer.Size = New System.Drawing.Size(74, 20)
        Me.CheckRefeshTimer.TabIndex = 15
        Me.CheckRefeshTimer.Text = "Active"
        Me.CheckRefeshTimer.UseVisualStyleBackColor = True
        '
        'Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(1073, 564)
        Me.Controls.Add(Me.CheckRefeshTimer)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtRefreshtimer)
        Me.Controls.Add(Me.btnUsersTable)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnNewConnectionSettings)
        Me.Controls.Add(Me.btnSavesettings)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.txtdatabaseName)
        Me.Controls.Add(Me.lblusernameSQL)
        Me.Controls.Add(Me.txtUsernameSQL)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.ComboBox1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Settings"
        Me.Text = "Setting"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents txtUsernameSQL As TextBox
    Friend WithEvents lblusernameSQL As Label
    Friend WithEvents btnConnect As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtdatabaseName As TextBox
    Friend WithEvents btnSavesettings As Button
    Friend WithEvents btnNewConnectionSettings As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents btnUsersTable As Button
    Friend WithEvents txtRefreshtimer As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents CheckRefeshTimer As CheckBox
End Class
