<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewUserForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewUserForm))
        Me.txtnewusername = New System.Windows.Forms.TextBox()
        Me.txtNewpassword1 = New System.Windows.Forms.TextBox()
        Me.txtnewpassword2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnRegister = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtnewusername
        '
        Me.txtnewusername.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtnewusername.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtnewusername.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtnewusername.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnewusername.ForeColor = System.Drawing.Color.Black
        Me.txtnewusername.Location = New System.Drawing.Point(352, 98)
        Me.txtnewusername.Name = "txtnewusername"
        Me.txtnewusername.Size = New System.Drawing.Size(366, 23)
        Me.txtnewusername.TabIndex = 0
        '
        'txtNewpassword1
        '
        Me.txtNewpassword1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtNewpassword1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtNewpassword1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewpassword1.ForeColor = System.Drawing.Color.Black
        Me.txtNewpassword1.Location = New System.Drawing.Point(352, 198)
        Me.txtNewpassword1.Name = "txtNewpassword1"
        Me.txtNewpassword1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNewpassword1.Size = New System.Drawing.Size(366, 23)
        Me.txtNewpassword1.TabIndex = 1
        Me.txtNewpassword1.UseSystemPasswordChar = True
        '
        'txtnewpassword2
        '
        Me.txtnewpassword2.AcceptsReturn = True
        Me.txtnewpassword2.AcceptsTab = True
        Me.txtnewpassword2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtnewpassword2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtnewpassword2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnewpassword2.ForeColor = System.Drawing.Color.Black
        Me.txtnewpassword2.Location = New System.Drawing.Point(352, 287)
        Me.txtnewpassword2.Name = "txtnewpassword2"
        Me.txtnewpassword2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtnewpassword2.Size = New System.Drawing.Size(366, 23)
        Me.txtnewpassword2.TabIndex = 2
        Me.txtnewpassword2.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(242, 101)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 22)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Username"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(243, 201)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 22)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Password"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(164, 287)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(182, 22)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Confirm Password"
        '
        'btnRegister
        '
        Me.btnRegister.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRegister.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRegister.ForeColor = System.Drawing.Color.White
        Me.btnRegister.Location = New System.Drawing.Point(579, 345)
        Me.btnRegister.Name = "btnRegister"
        Me.btnRegister.Size = New System.Drawing.Size(94, 36)
        Me.btnRegister.TabIndex = 7
        Me.btnRegister.Text = "Register"
        Me.btnRegister.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExit.ForeColor = System.Drawing.Color.White
        Me.btnExit.Location = New System.Drawing.Point(342, 343)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 38)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label3.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(281, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(222, 31)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "CREATE PROFILE"
        '
        'NewUserForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnRegister)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtnewpassword2)
        Me.Controls.Add(Me.txtNewpassword1)
        Me.Controls.Add(Me.txtnewusername)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "NewUserForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New User Registration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtnewusername As TextBox
    Friend WithEvents txtNewpassword1 As TextBox
    Friend WithEvents txtnewpassword2 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents btnRegister As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents Label3 As Label
End Class
