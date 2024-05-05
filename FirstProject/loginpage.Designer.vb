<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class LOGINPAGE
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LOGINPAGE))
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.USERNAME = New System.Windows.Forms.Label()
        Me.PASSWORD = New System.Windows.Forms.Label()
        Me.btnLogIn = New System.Windows.Forms.Button()
        Me.btnEXIT = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnChangePassword = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtUserName
        '
        Me.txtUserName.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        resources.ApplyResources(Me.txtUserName, "txtUserName")
        Me.txtUserName.ForeColor = System.Drawing.Color.White
        Me.txtUserName.Name = "txtUserName"
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.None
        resources.ApplyResources(Me.txtPassword, "txtPassword")
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'USERNAME
        '
        resources.ApplyResources(Me.USERNAME, "USERNAME")
        Me.USERNAME.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.USERNAME.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.USERNAME.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.USERNAME.Name = "USERNAME"
        '
        'PASSWORD
        '
        resources.ApplyResources(Me.PASSWORD, "PASSWORD")
        Me.PASSWORD.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.PASSWORD.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.PASSWORD.Name = "PASSWORD"
        '
        'btnLogIn
        '
        Me.btnLogIn.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        resources.ApplyResources(Me.btnLogIn, "btnLogIn")
        Me.btnLogIn.ForeColor = System.Drawing.Color.White
        Me.btnLogIn.Name = "btnLogIn"
        Me.btnLogIn.UseVisualStyleBackColor = False
        '
        'btnEXIT
        '
        Me.btnEXIT.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        resources.ApplyResources(Me.btnEXIT, "btnEXIT")
        Me.btnEXIT.ForeColor = System.Drawing.Color.White
        Me.btnEXIT.Name = "btnEXIT"
        Me.btnEXIT.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        resources.ApplyResources(Me.btnClose, "btnClose")
        Me.btnClose.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.Name = "btnClose"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Transparent
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'btnChangePassword
        '
        Me.btnChangePassword.BackColor = System.Drawing.Color.Transparent
        resources.ApplyResources(Me.btnChangePassword, "btnChangePassword")
        Me.btnChangePassword.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnChangePassword.Name = "btnChangePassword"
        Me.btnChangePassword.UseVisualStyleBackColor = False
        '
        'LOGINPAGE
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.Controls.Add(Me.btnChangePassword)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnEXIT)
        Me.Controls.Add(Me.btnLogIn)
        Me.Controls.Add(Me.PASSWORD)
        Me.Controls.Add(Me.USERNAME)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.txtUserName)
        Me.DoubleBuffered = True
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LOGINPAGE"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtPassword As TextBox
    Friend WithEvents USERNAME As Label
    Friend WithEvents PASSWORD As Label
    Friend WithEvents btnLogIn As Button
    Friend WithEvents btnEXIT As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents btnChangePassword As Button
    Public WithEvents txtUserName As TextBox
End Class
