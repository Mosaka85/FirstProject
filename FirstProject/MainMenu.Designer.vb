<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainMenu
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu))
        Me.btnCreatletable = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.cmbTables = New System.Windows.Forms.ComboBox()
        Me.btnShowtable = New System.Windows.Forms.Button()
        Me.btnExportData = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ComboTableOrView = New System.Windows.Forms.ComboBox()
        Me.lbltableorView = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbFilterColumnsnames = New System.Windows.Forms.ComboBox()
        Me.ComboOperators = New System.Windows.Forms.ComboBox()
        Me.txtWhereConditiontext = New System.Windows.Forms.TextBox()
        Me.txtShowmecode = New System.Windows.Forms.TextBox()
        Me.txtSQLtablename = New System.Windows.Forms.TextBox()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtNewSQLTable = New System.Windows.Forms.Label()
        Me.txtFilter = New System.Windows.Forms.Label()
        Me.btnLogOut = New System.Windows.Forms.Button()
        Me.cmbSELECTsheet = New System.Windows.Forms.ComboBox()
        Me.lblSELECTWORKSHEET = New System.Windows.Forms.Label()
        Me.btnCreatedatabase = New System.Windows.Forms.Button()
        Me.txtdatabasename = New System.Windows.Forms.TextBox()
        Me.btnCreatedabase = New System.Windows.Forms.Button()
        Me.lblCreatedatabase = New System.Windows.Forms.Label()
        Me.txtDabasenameSHOW = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnRefreshgrid = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnDataType = New System.Windows.Forms.Button()
        Me.txtNumberofRow = New System.Windows.Forms.TextBox()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.lblTimer = New System.Windows.Forms.Label()
        Me.btnCustumCodeSQL = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCreatletable
        '
        Me.btnCreatletable.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCreatletable.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCreatletable.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreatletable.Location = New System.Drawing.Point(13, 1000)
        Me.btnCreatletable.Name = "btnCreatletable"
        Me.btnCreatletable.Size = New System.Drawing.Size(104, 64)
        Me.btnCreatletable.TabIndex = 0
        Me.btnCreatletable.Text = "Import Excel to SQLTable"
        Me.btnCreatletable.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnClose.Location = New System.Drawing.Point(1841, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(45, 26)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "X"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeight = 29
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.GridColor = System.Drawing.Color.DimGray
        Me.DataGridView1.Location = New System.Drawing.Point(396, 93)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 51
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 9.0!)
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.RowTemplate.Height = 30
        Me.DataGridView1.Size = New System.Drawing.Size(1452, 693)
        Me.DataGridView1.TabIndex = 0
        '
        'cmbTables
        '
        Me.cmbTables.BackColor = System.Drawing.Color.Gray
        Me.cmbTables.DropDownHeight = 200
        Me.cmbTables.DropDownWidth = 269
        Me.cmbTables.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbTables.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbTables.ForeColor = System.Drawing.Color.Black
        Me.cmbTables.IntegralHeight = False
        Me.cmbTables.Location = New System.Drawing.Point(39, 150)
        Me.cmbTables.Name = "cmbTables"
        Me.cmbTables.Size = New System.Drawing.Size(269, 27)
        Me.cmbTables.TabIndex = 1
        '
        'btnShowtable
        '
        Me.btnShowtable.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnShowtable.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnShowtable.Location = New System.Drawing.Point(39, 312)
        Me.btnShowtable.Name = "btnShowtable"
        Me.btnShowtable.Size = New System.Drawing.Size(75, 23)
        Me.btnShowtable.TabIndex = 11
        Me.btnShowtable.Text = "VIEW"
        Me.btnShowtable.UseVisualStyleBackColor = False
        '
        'btnExportData
        '
        Me.btnExportData.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnExportData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExportData.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportData.Location = New System.Drawing.Point(15, 942)
        Me.btnExportData.Name = "btnExportData"
        Me.btnExportData.Size = New System.Drawing.Size(102, 52)
        Me.btnExportData.TabIndex = 12
        Me.btnExportData.Text = "Export Grid to Excel"
        Me.btnExportData.UseVisualStyleBackColor = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.ProgressBar1.ForeColor = System.Drawing.Color.Gray
        Me.ProgressBar1.Location = New System.Drawing.Point(254, 1041)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(216, 23)
        Me.ProgressBar1.TabIndex = 14
        '
        'ComboTableOrView
        '
        Me.ComboTableOrView.BackColor = System.Drawing.Color.DimGray
        Me.ComboTableOrView.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ComboTableOrView.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboTableOrView.ForeColor = System.Drawing.Color.Transparent
        Me.ComboTableOrView.FormattingEnabled = True
        Me.ComboTableOrView.Location = New System.Drawing.Point(39, 93)
        Me.ComboTableOrView.Name = "ComboTableOrView"
        Me.ComboTableOrView.Size = New System.Drawing.Size(269, 27)
        Me.ComboTableOrView.TabIndex = 15
        '
        'lbltableorView
        '
        Me.lbltableorView.AutoSize = True
        Me.lbltableorView.BackColor = System.Drawing.Color.Transparent
        Me.lbltableorView.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltableorView.ForeColor = System.Drawing.Color.Yellow
        Me.lbltableorView.Location = New System.Drawing.Point(35, 125)
        Me.lbltableorView.Name = "lbltableorView"
        Me.lbltableorView.Size = New System.Drawing.Size(210, 19)
        Me.lbltableorView.TabIndex = 16
        Me.lbltableorView.Text = "SELECT TABLE OR VIEW"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.Yellow
        Me.Label1.Location = New System.Drawing.Point(36, 67)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 19)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "SELECT SCHEMA"
        '
        'cmbFilterColumnsnames
        '
        Me.cmbFilterColumnsnames.BackColor = System.Drawing.Color.Gray
        Me.cmbFilterColumnsnames.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.cmbFilterColumnsnames.FormattingEnabled = True
        Me.cmbFilterColumnsnames.Location = New System.Drawing.Point(41, 234)
        Me.cmbFilterColumnsnames.Name = "cmbFilterColumnsnames"
        Me.cmbFilterColumnsnames.Size = New System.Drawing.Size(122, 27)
        Me.cmbFilterColumnsnames.TabIndex = 18
        '
        'ComboOperators
        '
        Me.ComboOperators.BackColor = System.Drawing.Color.Gray
        Me.ComboOperators.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.ComboOperators.FormattingEnabled = True
        Me.ComboOperators.Location = New System.Drawing.Point(168, 234)
        Me.ComboOperators.Name = "ComboOperators"
        Me.ComboOperators.Size = New System.Drawing.Size(112, 27)
        Me.ComboOperators.TabIndex = 19
        '
        'txtWhereConditiontext
        '
        Me.txtWhereConditiontext.BackColor = System.Drawing.Color.Gray
        Me.txtWhereConditiontext.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.txtWhereConditiontext.Location = New System.Drawing.Point(40, 276)
        Me.txtWhereConditiontext.Name = "txtWhereConditiontext"
        Me.txtWhereConditiontext.Size = New System.Drawing.Size(240, 30)
        Me.txtWhereConditiontext.TabIndex = 20
        '
        'txtShowmecode
        '
        Me.txtShowmecode.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txtShowmecode.ForeColor = System.Drawing.Color.White
        Me.txtShowmecode.Location = New System.Drawing.Point(31, 3)
        Me.txtShowmecode.Multiline = True
        Me.txtShowmecode.Name = "txtShowmecode"
        Me.txtShowmecode.Size = New System.Drawing.Size(307, 61)
        Me.txtShowmecode.TabIndex = 21
        '
        'txtSQLtablename
        '
        Me.txtSQLtablename.BackColor = System.Drawing.Color.Gray
        Me.txtSQLtablename.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.txtSQLtablename.Location = New System.Drawing.Point(123, 993)
        Me.txtSQLtablename.Name = "txtSQLtablename"
        Me.txtSQLtablename.Size = New System.Drawing.Size(231, 30)
        Me.txtSQLtablename.TabIndex = 22
        '
        'btnImport
        '
        Me.btnImport.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnImport.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImport.Location = New System.Drawing.Point(123, 1029)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(125, 35)
        Me.btnImport.TabIndex = 23
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = False
        '
        'txtNewSQLTable
        '
        Me.txtNewSQLTable.AutoSize = True
        Me.txtNewSQLTable.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.txtNewSQLTable.ForeColor = System.Drawing.Color.Yellow
        Me.txtNewSQLTable.Location = New System.Drawing.Point(123, 971)
        Me.txtNewSQLTable.Name = "txtNewSQLTable"
        Me.txtNewSQLTable.Size = New System.Drawing.Size(105, 19)
        Me.txtNewSQLTable.TabIndex = 24
        Me.txtNewSQLTable.Text = "Table Name:"
        '
        'txtFilter
        '
        Me.txtFilter.AutoSize = True
        Me.txtFilter.BackColor = System.Drawing.Color.Transparent
        Me.txtFilter.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.txtFilter.ForeColor = System.Drawing.Color.Yellow
        Me.txtFilter.Location = New System.Drawing.Point(38, 207)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(49, 19)
        Me.txtFilter.TabIndex = 25
        Me.txtFilter.Text = "Filter"
        '
        'btnLogOut
        '
        Me.btnLogOut.BackColor = System.Drawing.Color.Yellow
        Me.btnLogOut.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnLogOut.Location = New System.Drawing.Point(1700, 8)
        Me.btnLogOut.Name = "btnLogOut"
        Me.btnLogOut.Size = New System.Drawing.Size(79, 33)
        Me.btnLogOut.TabIndex = 26
        Me.btnLogOut.Text = "LogOut"
        Me.btnLogOut.UseVisualStyleBackColor = False
        '
        'cmbSELECTsheet
        '
        Me.cmbSELECTsheet.BackColor = System.Drawing.Color.Gray
        Me.cmbSELECTsheet.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.cmbSELECTsheet.FormattingEnabled = True
        Me.cmbSELECTsheet.Location = New System.Drawing.Point(123, 942)
        Me.cmbSELECTsheet.Name = "cmbSELECTsheet"
        Me.cmbSELECTsheet.Size = New System.Drawing.Size(231, 27)
        Me.cmbSELECTsheet.TabIndex = 27
        '
        'lblSELECTWORKSHEET
        '
        Me.lblSELECTWORKSHEET.AutoSize = True
        Me.lblSELECTWORKSHEET.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.lblSELECTWORKSHEET.ForeColor = System.Drawing.Color.Yellow
        Me.lblSELECTWORKSHEET.Location = New System.Drawing.Point(120, 922)
        Me.lblSELECTWORKSHEET.Name = "lblSELECTWORKSHEET"
        Me.lblSELECTWORKSHEET.Size = New System.Drawing.Size(150, 19)
        Me.lblSELECTWORKSHEET.TabIndex = 28
        Me.lblSELECTWORKSHEET.Text = "Select WorkSheet"
        '
        'btnCreatedatabase
        '
        Me.btnCreatedatabase.BackColor = System.Drawing.Color.DarkKhaki
        Me.btnCreatedatabase.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCreatedatabase.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreatedatabase.Location = New System.Drawing.Point(13, 598)
        Me.btnCreatedatabase.Name = "btnCreatedatabase"
        Me.btnCreatedatabase.Size = New System.Drawing.Size(102, 59)
        Me.btnCreatedatabase.TabIndex = 29
        Me.btnCreatedatabase.Text = "Create Database"
        Me.btnCreatedatabase.UseVisualStyleBackColor = False
        '
        'txtdatabasename
        '
        Me.txtdatabasename.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.txtdatabasename.Location = New System.Drawing.Point(121, 853)
        Me.txtdatabasename.Name = "txtdatabasename"
        Me.txtdatabasename.Size = New System.Drawing.Size(227, 27)
        Me.txtdatabasename.TabIndex = 30
        '
        'btnCreatedabase
        '
        Me.btnCreatedabase.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCreatedabase.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCreatedabase.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreatedabase.ForeColor = System.Drawing.Color.Black
        Me.btnCreatedabase.Location = New System.Drawing.Point(127, 886)
        Me.btnCreatedabase.Name = "btnCreatedabase"
        Me.btnCreatedabase.Size = New System.Drawing.Size(75, 33)
        Me.btnCreatedabase.TabIndex = 31
        Me.btnCreatedabase.Text = "Create"
        Me.btnCreatedabase.UseVisualStyleBackColor = False
        '
        'lblCreatedatabase
        '
        Me.lblCreatedatabase.AutoSize = True
        Me.lblCreatedatabase.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold)
        Me.lblCreatedatabase.ForeColor = System.Drawing.Color.Yellow
        Me.lblCreatedatabase.Location = New System.Drawing.Point(121, 833)
        Me.lblCreatedatabase.Name = "lblCreatedatabase"
        Me.lblCreatedatabase.Size = New System.Drawing.Size(131, 19)
        Me.lblCreatedatabase.TabIndex = 32
        Me.lblCreatedatabase.Text = "Database Name"
        '
        'txtDabasenameSHOW
        '
        Me.txtDabasenameSHOW.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.txtDabasenameSHOW.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtDabasenameSHOW.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDabasenameSHOW.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.txtDabasenameSHOW.Location = New System.Drawing.Point(968, 1049)
        Me.txtDabasenameSHOW.Name = "txtDabasenameSHOW"
        Me.txtDabasenameSHOW.Size = New System.Drawing.Size(166, 15)
        Me.txtDabasenameSHOW.TabIndex = 33
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Yellow
        Me.Label2.Location = New System.Drawing.Point(38, 183)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(133, 19)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Numer Of Rows"
        '
        'btnRefreshgrid
        '
        Me.btnRefreshgrid.BackgroundImage = CType(resources.GetObject("btnRefreshgrid.BackgroundImage"), System.Drawing.Image)
        Me.btnRefreshgrid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnRefreshgrid.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRefreshgrid.ForeColor = System.Drawing.Color.Transparent
        Me.btnRefreshgrid.Image = CType(resources.GetObject("btnRefreshgrid.Image"), System.Drawing.Image)
        Me.btnRefreshgrid.Location = New System.Drawing.Point(123, 312)
        Me.btnRefreshgrid.Name = "btnRefreshgrid"
        Me.btnRefreshgrid.Size = New System.Drawing.Size(28, 24)
        Me.btnRefreshgrid.TabIndex = 36
        Me.btnRefreshgrid.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'btnDataType
        '
        Me.btnDataType.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDataType.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataType.ForeColor = System.Drawing.Color.Black
        Me.btnDataType.Location = New System.Drawing.Point(1681, 922)
        Me.btnDataType.Name = "btnDataType"
        Me.btnDataType.Size = New System.Drawing.Size(79, 30)
        Me.btnDataType.TabIndex = 37
        Me.btnDataType.Text = "Import"
        Me.btnDataType.UseVisualStyleBackColor = False
        '
        'txtNumberofRow
        '
        Me.txtNumberofRow.Location = New System.Drawing.Point(180, 180)
        Me.txtNumberofRow.Name = "txtNumberofRow"
        Me.txtNumberofRow.Size = New System.Drawing.Size(100, 22)
        Me.txtNumberofRow.TabIndex = 38
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'lblTimer
        '
        Me.lblTimer.AutoSize = True
        Me.lblTimer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTimer.ForeColor = System.Drawing.Color.LawnGreen
        Me.lblTimer.Location = New System.Drawing.Point(39, 338)
        Me.lblTimer.Name = "lblTimer"
        Me.lblTimer.Size = New System.Drawing.Size(46, 19)
        Me.lblTimer.TabIndex = 39
        Me.lblTimer.Text = "Timer"
        '
        'btnCustumCodeSQL
        '
        Me.btnCustumCodeSQL.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCustumCodeSQL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCustumCodeSQL.Location = New System.Drawing.Point(344, 34)
        Me.btnCustumCodeSQL.Name = "btnCustumCodeSQL"
        Me.btnCustumCodeSQL.Size = New System.Drawing.Size(75, 30)
        Me.btnCustumCodeSQL.TabIndex = 40
        Me.btnCustumCodeSQL.Text = "Run"
        Me.btnCustumCodeSQL.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button2.BackgroundImage = CType(resources.GetObject("Button2.BackgroundImage"), System.Drawing.Image)
        Me.Button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(1635, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(59, 56)
        Me.Button2.TabIndex = 41
        Me.Button2.UseVisualStyleBackColor = False
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(51, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1920, 1080)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnCustumCodeSQL)
        Me.Controls.Add(Me.lblTimer)
        Me.Controls.Add(Me.txtNumberofRow)
        Me.Controls.Add(Me.btnDataType)
        Me.Controls.Add(Me.btnRefreshgrid)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDabasenameSHOW)
        Me.Controls.Add(Me.lblCreatedatabase)
        Me.Controls.Add(Me.btnCreatedabase)
        Me.Controls.Add(Me.txtdatabasename)
        Me.Controls.Add(Me.lblSELECTWORKSHEET)
        Me.Controls.Add(Me.cmbSELECTsheet)
        Me.Controls.Add(Me.btnLogOut)
        Me.Controls.Add(Me.txtFilter)
        Me.Controls.Add(Me.txtNewSQLTable)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.txtSQLtablename)
        Me.Controls.Add(Me.txtShowmecode)
        Me.Controls.Add(Me.txtWhereConditiontext)
        Me.Controls.Add(Me.ComboOperators)
        Me.Controls.Add(Me.cmbFilterColumnsnames)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbltableorView)
        Me.Controls.Add(Me.ComboTableOrView)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnExportData)
        Me.Controls.Add(Me.btnShowtable)
        Me.Controls.Add(Me.cmbTables)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCreatletable)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "MainMenu"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCreatletable As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents cmbTables As ComboBox
    Friend WithEvents btnShowtable As Button
    Friend WithEvents btnExportData As Button
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents ComboTableOrView As ComboBox
    Friend WithEvents lbltableorView As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbFilterColumnsnames As ComboBox
    Friend WithEvents ComboOperators As ComboBox
    Friend WithEvents txtShowmecode As TextBox
    Friend WithEvents txtNewSQLTable As Label
    Friend WithEvents txtFilter As Label
    Friend WithEvents btnLogOut As Button
    Friend WithEvents cmbSELECTsheet As ComboBox
    Friend WithEvents lblSELECTWORKSHEET As Label
    Friend WithEvents btnCreatedatabase As Button
    Friend WithEvents txtdatabasename As TextBox
    Friend WithEvents btnCreatedabase As Button
    Friend WithEvents lblCreatedatabase As Label
    Friend WithEvents txtDabasenameSHOW As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnRefreshgrid As Button
    Public WithEvents Timer1 As Timer
    Public WithEvents DataGridView1 As DataGridView
    Public WithEvents btnImport As Button
    Friend WithEvents btnDataType As Button
    Public WithEvents txtWhereConditiontext As TextBox
    Public WithEvents txtSQLtablename As TextBox
    Friend WithEvents txtNumberofRow As TextBox
    Friend WithEvents PrintPreviewDialog1 As PrintPreviewDialog
    Friend WithEvents lblTimer As Label
    Friend WithEvents btnCustumCodeSQL As Button
    Friend WithEvents Button2 As Button
End Class
