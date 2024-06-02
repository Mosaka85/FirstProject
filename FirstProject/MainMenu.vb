Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports DataTable = System.Data.DataTable
Imports Workbook = Microsoft.Office.Interop.Excel.Workbook
Imports Application = Microsoft.Office.Interop.Excel.Application
Imports System.Text
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Microsoft.Reporting.WebForms
Imports System.Drawing.Printing
Imports Microsoft.Reporting.WinForms
Imports WinFormsReport = Microsoft.Reporting.WinForms
Imports WebFormsReport = Microsoft.Reporting.WebForms
Imports System.Web
Imports ProcessingMode = Microsoft.Reporting.WebForms.ProcessingMode
Imports Warning = Microsoft.Reporting.WebForms.Warning
Imports ReportDataSource = Microsoft.Reporting.WebForms.ReportDataSource
Imports System.Drawing
Imports Font = System.Drawing.Font

Public Class MainMenu
    Dim configFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "settings.xml")
    Dim xmlDoc As New XmlDocument()
    Dim connectionString As String
    Private bgWorker As Object
    Dim excelFilePath As String
    Dim dt As String
    Dim tablename As String
    Dim selectedSheet As String
    Dim RefreshTime As String
    Private isTimerRunning As Boolean = False
    Private countdown As Integer
    Dim refreshActiveIndicator As Integer
    Private Sub menuL_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        btnCreatedatabase.Visible = False
        If LOGINPAGE.txtUserName.Text = "Mosaka" Then
            btnCreatedatabase.Visible = True
        End If
        ProgressBar1.Visible = False
        btnImport.Visible = False
        txtSQLtablename.Visible = False
        txtNewSQLTable.Visible = False
        lblSELECTWORKSHEET.Visible = False
        cmbSELECTsheet.Visible = False
        txtdatabasename.Visible = False
        btnCreatedabase.Visible = False
        lblCreatedatabase.Visible = False
        btnDataType.Visible = False
        Try
            xmlDoc.Load(configFile)

            ' Use XmlNamespaceManager to handle default namespaces
            Dim xmlNsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            xmlNsmgr.AddNamespace("default", xmlDoc.DocumentElement.NamespaceURI)

            ' Select nodes with namespace manager
            Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='DataSource']", xmlNsmgr)
            Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='Username']", xmlNsmgr)
            Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='Password']", xmlNsmgr)
            Dim databaseNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='InitialCatalog']", xmlNsmgr)
            Dim refreshTimeIntervalNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='RefreshTime']", xmlNsmgr)
            Dim refreshActiveTimerNode As XmlNode = xmlDoc.SelectSingleNode("//default:appSettings/default:add[@key='CheckRefeshTimer']", xmlNsmgr)

            ' Check for nulls and assign values
            refreshActiveIndicator = If(refreshActiveTimerNode IsNot Nothing AndAlso Integer.TryParse(refreshActiveTimerNode.Attributes("value").Value, Nothing), Integer.Parse(refreshActiveTimerNode.Attributes("value").Value), 0) ' Default to 0 if not found
            RefreshTime = If(refreshTimeIntervalNode IsNot Nothing, Integer.Parse(refreshTimeIntervalNode.Attributes("value").Value), 20) ' Default to 20 seconds if not found
            Dim dataSource As String = If(dataSourceNode IsNot Nothing, dataSourceNode.Attributes("value").Value, "")
            Dim username As String = If(usernameNode IsNot Nothing, usernameNode.Attributes("value").Value, "")
            Dim password As String = If(passwordNode IsNot Nothing, passwordNode.Attributes("value").Value, "")
            Dim database As String = If(databaseNode IsNot Nothing, databaseNode.Attributes("value").Value, "")

            txtDabasenameSHOW.Text = database
            connectionString = $"Data Source={dataSource};Initial Catalog={database};User ID={username};Password={password};"
            lblTimer.Text = $"Time until refresh: {RefreshTime} seconds"
            ' Start timer only if refreshActiveIndicator is 1
            If refreshActiveIndicator = "1" Then
                lblTimer.Visible = True
            Else
                lblTimer.Visible = False
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading settings: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillCombo(ComboTableOrView)
    End Sub
    Private Sub FillCombo(comboBox As ComboBox)
        comboBox.Items.Clear()
        comboBox.Items.Add("Tables")
        comboBox.Items.Add("Views")
        comboBox.Sorted = True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If countdown > 0 AndAlso refreshActiveIndicator = "1" Then
            countdown -= 1
            lblTimer.Text = "Time until refresh: " & countdown & " seconds"
        Else
            RefreshData()
            countdown = RefreshTime ' Reset countdown
        End If
    End Sub
    Private Sub btnShowtable_Click(sender As Object, e As EventArgs) Handles btnShowtable.Click
        If refreshActiveIndicator = "1" Then
            Timer1.Start()
            countdown = RefreshTime
        End If

        Dim numberofrows As String = ""
        Dim parameterName As String = ""
        Dim parameterValue As Object = Nothing

        If txtNumberofRow.Text <> "" AndAlso IsNumeric(txtNumberofRow.Text) Then
            numberofrows = " TOP " & txtNumberofRow.Text
        End If

        Try
            If cmbTables.SelectedItem IsNot Nothing Then
                Dim selectedTable As String = cmbTables.SelectedItem.ToString()
                Dim whereClause As String = ""

                If Not String.IsNullOrWhiteSpace(txtWhereConditiontext.Text) AndAlso
           cmbFilterColumnsnames.SelectedItem IsNot Nothing AndAlso
           ComboOperators.SelectedItem IsNot Nothing Then

                    Dim selectedColumn As String = cmbFilterColumnsnames.SelectedItem.ToString()
                    Dim selectedOperator As String = ComboOperators.SelectedItem.ToString()
                    Dim value As String = txtWhereConditiontext.Text

                    ' Use parameterized query to prevent SQL injection
                    parameterName = $"@{selectedColumn.Replace("[", "").Replace("]", "")}"
                    parameterValue = If(IsNumeric(value), Convert.ToInt32(value), value)

                    whereClause = $" WHERE [{selectedColumn}] {selectedOperator} {parameterName}"
                End If

                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    Dim query As String = $"SELECT {numberofrows} * FROM {selectedTable} {whereClause}"
                    txtShowmecode.Text = query

                    Using command As New SqlCommand(query, connection)
                        ' Add parameters if any
                        If Not String.IsNullOrWhiteSpace(whereClause) Then
                            command.Parameters.AddWithValue(parameterName, parameterValue)
                        End If

                        Dim adapter As New SqlDataAdapter(command)
                        Dim dataSet As New DataSet()
                        adapter.Fill(dataSet)

                        If dataSet.Tables(0).Rows.Count = 0 Then
                            MessageBox.Show("The query returned no results.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            DataGridView1.DataSource = dataSet.Tables(0)
                        End If
                    End Using
                End Using
            Else
                customMsgBoxF.Show()
                customMsgBoxF.txtMsgSucess.Text = "Please select a table first."
            End If

        Catch ex As SqlException
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = "Database Error: " & ex.Message
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = "An error occurred: " & ex.Message
        End Try
    End Sub


    Private oExcel As Microsoft.Office.Interop.Excel.Application
    Private oBook As Microsoft.Office.Interop.Excel.Workbook
    Private oSheet As Microsoft.Office.Interop.Excel.Worksheet
    Private dgvData As Object
    Private headerFont As System.Drawing.Font

    Private Sub BtnExportData_Click(sender As Object, e As EventArgs) Handles btnExportData.Click
        cmbSELECTsheet.Visible = False
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        saveFileDialog.FilterIndex = 1
        saveFileDialog.RestoreDirectory = True
        ProgressBar1.Visible = True
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            Try
                oBook = excelApp.Workbooks.Add()
                oSheet = CType(oBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    oSheet.Cells(1, j + 1).Value = DataGridView1.Columns(j).HeaderText
                Next
                ProgressBar1.Maximum = DataGridView1.Rows.Count
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    ProgressBar1.Value = i + 1

                    For j As Integer = 0 To DataGridView1.Columns.Count - 1
                        oSheet.Cells(i + 2, j + 1).Value = If(DataGridView1.Rows(i).Cells(j).Value IsNot Nothing, DataGridView1.Rows(i).Cells(j).Value.ToString(), "")
                    Next
                Next
                oBook.SaveAs(saveFileDialog.FileName)
                MessageBox.Show("Data exported successfully To Excel.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = "Error exporting data To Excel:  " & ex.Message
            Finally
                oBook?.Close(SaveChanges:=False)
                oExcel?.Quit()
                ProgressBar1.Visible = False
                If excelApp IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                    excelApp = Nothing
                End If
            End Try
        End If
    End Sub
    Private Sub ComboTableOrView_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboTableOrView.SelectedIndexChanged
        cmbTables.Items.Clear()
        cmbTables.Text = ""
        txtWhereConditiontext.Text = ""
        txtWhereConditiontext.Text = ""
        txtSQLtablename.Text = ""
        ComboOperators.Text = ""

        cmbFilterColumnsnames.Items.Clear()
        cmbFilterColumnsnames.Text = ""
        If ComboTableOrView.SelectedItem IsNot Nothing Then
            Dim selectedItem As String = ComboTableOrView.SelectedItem.ToString()
            Dim showViews As Boolean = (selectedItem = "Views")
            PopulateTablesComboBox(showViews)
        End If
    End Sub
    Private Sub PopulateTablesComboBox(showViews As Boolean)
        cmbTables.Items.Clear()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim schemaType As String = If(showViews, "VIEWS", "TABLES")
                Dim schemaResult As DataTable = connection.GetSchema(schemaType)

                If schemaResult IsNot Nothing AndAlso schemaResult.Rows.Count > 0 Then
                    For Each row As DataRow In schemaResult.Rows
                        cmbTables.Items.Add(row("TABLE_NAME").ToString())
                    Next
                Else
                    MessageBox.Show($"No {(If(showViews, "views", "tables"))} found in the database.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        Catch ex As SqlException
            MessageBox.Show("Database Error: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        cmbTables.Sorted = True ' Enable sorting
    End Sub
    Private Sub PopulateColumnsComboBox(tableName As String)
        cmbFilterColumnsnames.Items.Clear()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim columnsSchema As DataTable = connection.GetSchema("Columns", New String() {Nothing, Nothing, tableName})
                If columnsSchema IsNot Nothing AndAlso columnsSchema.Rows.Count > 0 Then
                    For Each row As DataRow In columnsSchema.Rows
                        Dim columnName As String = row("COLUMN_NAME").ToString()

                        cmbFilterColumnsnames.Items.Add(columnName)
                    Next
                Else
                    ErrorPage.Show()
                    ErrorPage.txtErrorMassage.Text = $"No columns found for table '{tableName}'."
                End If
            End Using
        Catch ex As Exception
            ' Handle any exceptions
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = "Error: " & ex.Message
        End Try
    End Sub
    Private Sub FillOperatorsCombo(comboBox As ComboBox)
        comboBox.Items.Clear()
        comboBox.Items.Add("=")
        comboBox.Items.Add("<")
        comboBox.Items.Add(">")
        comboBox.Items.Add("<>")
        comboBox.Items.Add(">=")
        comboBox.Items.Add("<=")
        comboBox.Items.Add("LIKE")
        comboBox.Items.Add("IN")
        comboBox.Sorted = True
    End Sub

    Private Sub cmbTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTables.SelectedIndexChanged
        txtWhereConditiontext.Text = ""
        txtWhereConditiontext.Text = ""
        cmbFilterColumnsnames.Items.Clear()
        cmbFilterColumnsnames.Text = ""
        ComboOperators.Items.Clear()
        ComboOperators.Text = ""

        If cmbTables.SelectedItem IsNot Nothing Then
            Dim selectedTable As String = cmbTables.SelectedItem.ToString()
            PopulateColumnsComboBox(selectedTable)
        End If

        FillOperatorsCombo(ComboOperators)
    End Sub
    Private Sub btnCreatletable_Click(sender As Object, e As EventArgs) Handles btnCreatletable.Click
        Timer1.Stop()
        cmbSELECTsheet.Visible = True
        cmbSELECTsheet.Text = ""
        txtSQLtablename.Text = ""
        lblSELECTWORKSHEET.Visible = True
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls"
        openFileDialog1.Title = "Select an Excel File"
        If openFileDialog1.ShowDialog() = DialogResult.OK AndAlso File.Exists(openFileDialog1.FileName) Then
            Dim xlApp As Application = Nothing
            Dim xlWorkBook As Workbook = Nothing
            Dim xlWorkSheet As Worksheet = Nothing
            Try
                xlApp = New Application
                xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName)
                cmbSELECTsheet.Items.Clear()
                For Each xlSheet As Worksheet In xlWorkBook.Sheets
                    cmbSELECTsheet.Items.Add(xlSheet.Name)
                Next
                excelFilePath = openFileDialog1.FileName
            Catch ex As Exception
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = $"An error occurred opening Excel file: {ex.Message}"
            Finally
                If xlWorkBook IsNot Nothing Then xlWorkBook.Close()
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    ReleaseObject(xlApp)
                End If
            End Try
        Else
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = "The selected file was not found. Please try again."
        End If
    End Sub
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        btnDataType.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Try
            DataGridView1.DataSource = Nothing
            btnDataType.Visible = True
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            If txtSQLtablename.Text = "" Then
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = "Please Enter Table name before import"
                Return
            End If
            If cmbSELECTsheet.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(cmbSELECTsheet.SelectedItem.ToString()) Then
                MessageBox.Show("Please select a sheet before importing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim selectedSheet As String = cmbSELECTsheet.SelectedItem.ToString()
            Dim excelConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"

            Using excelConnection As New OleDb.OleDbConnection(excelConnectionString)
                excelConnection.Open()
                Dim dt As New DataTable()
                Dim query As String = "SELECT * FROM [" & selectedSheet & "$]"
                Using adapter As New OleDbDataAdapter(query, excelConnection)
                    adapter.FillSchema(dt, SchemaType.Source)
                End Using

                Dim dataTypeForm As New SelectDatatype()

                DataGridView1.Columns.Add("ColumnName", "Column Name")
                Dim comboBoxColumn As New DataGridViewComboBoxColumn()
                comboBoxColumn.HeaderText = "Data Type"
                comboBoxColumn.Name = "DataTypeColumn"
                comboBoxColumn.Items.AddRange("NVARCHAR(MAX)", "INT", "FLOAT", "BIT", "DATETIME2", "DECIMAL(18,6)")
                DataGridView1.Columns.Add(comboBoxColumn)


                For Each column As DataColumn In dt.Columns
                    DataGridView1.Rows.Add(column.ColumnName)
                Next

            End Using
            btnImport.Visible = False
            txtSQLtablename.Visible = False
            txtNewSQLTable.Visible = False
            lblSELECTWORKSHEET.Visible = False
            cmbSELECTsheet.Visible = False
        Catch ex As Exception
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = "Error Importing File: " & ex.Message
        End Try
    End Sub
    Private Sub cmbSELECTsheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSELECTsheet.SelectedIndexChanged
        btnImport.Visible = True
        txtSQLtablename.Visible = True
        txtNewSQLTable.Visible = True
        lblSELECTWORKSHEET.Visible = True
        cmbSELECTsheet.Visible = True
    End Sub
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Private Sub btnLogOut_Click(sender As Object, e As EventArgs) Handles btnLogOut.Click
        Me.Hide()
        LOGINPAGE.Show()
        Timer1.Stop()
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Hide()
        LOGINPAGE.Show()
    End Sub
    Private Sub btnCreatedabase_Click(sender As Object, e As EventArgs) Handles btnCreatedabase.Click
        If String.IsNullOrWhiteSpace(txtdatabasename.Text.Trim()) Then
            MessageBox.Show("Please specify a database name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        xmlDoc.Load(configFile)
        Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='DataSource']")
        Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Username']")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Password']")
        Dim databaseNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='InitialCatalog']")
        Dim dataSource As String = dataSourceNode.Attributes("value").Value
        Dim username As String = usernameNode.Attributes("value").Value
        Dim password As String = passwordNode.Attributes("value").Value
        Dim database As String = databaseNode.Attributes("value").Value
        Dim connectionString As String = $"Data Source={dataSource};Initial Catalog=master;User ID={username};Password={password};"
        Dim query As String = $"CREATE DATABASE {txtdatabasename.Text.Trim()}"
        txtShowmecode.Text = query
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                Try
                    connection.Open()
                    command.ExecuteNonQuery()
                    MessageBox.Show("Database created successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    ErrorPage.Show()
                    ErrorPage.txtErrorMassage.Text = "Error creating database: " & ex.Message
                End Try
            End Using
        End Using
    End Sub
    Private Sub btnCreatedatabase_Click(sender As Object, e As EventArgs) Handles btnCreatedatabase.Click
        txtdatabasename.Visible = True
        btnCreatedabase.Visible = True
        lblCreatedatabase.Visible = True
    End Sub
    Private Sub btnRefreshGrid_Click(sender As Object, e As EventArgs) Handles btnRefreshgrid.Click
        Dim numberofrows As String = ""
        If Not String.IsNullOrWhiteSpace(txtNumberofRow.Text) Then
            numberofrows = " TOP " & txtNumberofRow.Text
        End If
        Try
            If cmbTables.SelectedItem IsNot Nothing Then
                Dim selectedTable As String = cmbTables.SelectedItem.ToString()
                Dim whereClause As String = ""
                If txtWhereConditiontext.Text.Trim() = $"" Then
                    whereClause = ""
                ElseIf cmbFilterColumnsnames.SelectedItem IsNot Nothing AndAlso ComboOperators.SelectedItem IsNot Nothing Then
                    Dim selectedColumn As String = cmbFilterColumnsnames.SelectedItem.ToString()
                    Dim selectedOperator As String = ComboOperators.SelectedItem.ToString()
                    Dim value As String = txtWhereConditiontext.Text
                    If Not IsNumeric(value) Then

                        value = "'" & value & "'"
                    End If
                    whereClause = $" WHERE {selectedColumn} {selectedOperator} {value}"
                End If
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT " & numberofrows & " *  FROM  " & selectedTable & "  " & whereClause
                    txtShowmecode.Text = query
                    Dim adapter As New SqlDataAdapter(query, connection)
                    Dim dataSet As New DataSet()
                    adapter.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                End Using
            Else
                ErrorPage.Show()
                customMsgBoxF.txtMsgSucess.Text = "Please select a table first."
            End If
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = ex.Message
        End Try
        PopulateDataGridView()
    End Sub
    Private Sub PopulateDataGridView()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        Try
            If cmbTables.SelectedItem IsNot Nothing Then
                Dim selectedTable As String = cmbTables.SelectedItem.ToString()
                Dim whereClause As String = ""
                If txtWhereConditiontext.Text.Trim() <> "" Then
                    whereClause = " WHERE ..."
                End If
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT * FROM " & selectedTable & whereClause
                    Dim adapter As New SqlDataAdapter(query, connection)
                    Dim dataSet As New DataSet()
                    adapter.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                End Using
            Else
                MessageBox.Show("Please select a table first.")
            End If
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = ex.Message
        End Try
    End Sub
    Private Function GetSqlDataType(dataType As Type, columnName As String) As String
        ' Map .NET data types to SQL data types with enhancements
        Dim sqlDataTypes As Dictionary(Of Type, String) = New Dictionary(Of Type, String) From {
        {GetType(String), "NVARCHAR(MAX)"},
        {GetType(Integer), "INT"},
        {GetType(Double), "FLOAT"},
        {GetType(Boolean), "BIT"},
        {GetType(DateTime), "DATETIME2"},
        {GetType(Decimal), "DECIMAL(18,6)"}
    }
        If dataType IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnName) Then

            If sqlDataTypes.ContainsKey(dataType) Then
                Return sqlDataTypes(dataType)
            Else
                Return "NVARCHAR(MAX)"
            End If
        Else
            Return "NVARCHAR(MAX)"
        End If
    End Function
    Private Sub btnDataType_Click_1(sender As Object, e As EventArgs) Handles btnDataType.Click
        Dim tablename = txtSQLtablename.Text
        If cmbSELECTsheet.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(txtNewSQLTable.Text) Then
            MessageBox.Show("Please select a sheet and specify a SQL table name before importing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim excelConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim selectedSheet As String = cmbSELECTsheet.SelectedItem?.ToString()
        If String.IsNullOrEmpty(selectedSheet) Then
            MessageBox.Show("Please select a sheet before importing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        ProgressBar1.Visible = True
        ProgressBar1.Minimum = 0
        Dim dt As New DataTable()
        Using excelConnection As New OleDb.OleDbConnection(excelConnectionString)
            excelConnection.Open()
            Dim query As String = "SELECT * FROM [" & selectedSheet & "$]"
            Using adapter As New OleDbDataAdapter(query, excelConnection)
                adapter.Fill(dt)
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    MessageBox.Show("No data found in the selected sheet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                ProgressBar1.Maximum = dt.Rows.Count
            End Using
        End Using
        Dim createTableQuery As New StringBuilder($"CREATE TABLE [{tablename}] (")
        Dim columnsAdded As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow AndAlso row.Cells.Count >= 2 Then
                Dim columnName As String = row.Cells(0).Value?.ToString()
                Dim dataType As String = row.Cells(1).Value?.ToString()
                If Not String.IsNullOrWhiteSpace(columnName) Then
                    If String.IsNullOrWhiteSpace(dataType) Then
                        dataType = "nvarchar(max)"
                    End If
                    createTableQuery.Append($" [{columnName.Replace(" ", "_")}] {dataType},")
                    columnsAdded += 1
                End If
            End If
        Next
        If columnsAdded > 0 Then
            createTableQuery.Remove(createTableQuery.Length - 1, 1) ' Remove the last comma
        End If
        createTableQuery.Append(")")
        txtShowmecode.Text = createTableQuery.ToString()
        Console.WriteLine("Generated Query: " & createTableQuery.ToString())
        Try
            Using sqlConnection As New SqlConnection(connectionString)
                sqlConnection.Open()
                Using transaction As SqlTransaction = sqlConnection.BeginTransaction()
                    Using createTableCommand As New SqlCommand(createTableQuery.ToString(), sqlConnection, transaction)
                        createTableCommand.ExecuteNonQuery()
                    End Using
                    Using bulkCopy As New SqlBulkCopy(sqlConnection, SqlBulkCopyOptions.Default, transaction)
                        bulkCopy.DestinationTableName = tablename
                        AddHandler bulkCopy.SqlRowsCopied, Sub(senderObj, eventArgs)
                                                               ' Update progress bar value
                                                               ProgressBar1.Value = eventArgs.RowsCopied
                                                           End Sub
                        bulkCopy.WriteToServer(dt)
                    End Using
                    transaction.Commit()
                    MessageBox.Show($"Data imported successfully table name:[{tablename}].", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End Using
        Catch sqlEx As SqlException
            MessageBox.Show("Error executing database operations: " & sqlEx.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("An unexpected error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ProgressBar1.Visible = False
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
    End Sub
    Private Function ValidateUserInput() As Boolean
        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        If String.IsNullOrEmpty(selectedSheet) Then
            MessageBox.Show("Please select a sheet before importing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        If String.IsNullOrEmpty(txtSQLtablename.Text) Then
            MessageBox.Show("Please enter a name for the SQL Server table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
    Private Function TryGetExcelData() As DataTable
        cmbSELECTsheet.Visible = True
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls"
        openFileDialog1.Title = "Select an Excel File"
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim xlApp As Application = Nothing
            Dim xlWorkBook As Workbook = Nothing
            Dim xlWorkSheet As Worksheet = Nothing
            Try
                xlApp = New Application
                xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName)
                cmbSELECTsheet.Items.Clear()
                For Each xlSheet As Worksheet In xlWorkBook.Sheets
                    cmbSELECTsheet.Items.Add(xlSheet.Name)
                Next
                excelFilePath = openFileDialog1.FileName
            Catch ex As Exception
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = $"An error occurred opening Excel file: {ex.Message}"
            Finally
                If xlWorkBook IsNot Nothing Then xlWorkBook.Close()
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    ReleaseObject(xlApp)
                End If
            End Try
        End If
    End Function
    Private Sub PopulateDataGridView(dt As DataTable)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ColumnName", "Column Name")
        Dim dataTypeColumn As New DataGridViewComboBoxColumn()
        dataTypeColumn.HeaderText = "Data Type"
        dataTypeColumn.Name = "DataTypeColumn"
        dataTypeColumn.Items.AddRange("NVARCHAR(MAX)", "INT", "FLOAT", "BIT", "DATETIME2", "DECIMAL(18,6)")
        DataGridView1.Columns.Add(dataTypeColumn)
        For Each column As DataColumn In dt.Columns
            Dim rowIndex As Integer = DataGridView1.Rows.Add(column.ColumnName)
            Dim cell As DataGridViewComboBoxCell = CType(DataGridView1.Rows(rowIndex).Cells("DataTypeColumn"), DataGridViewComboBoxCell)
            Dim dataTypeIndex As Integer = dataTypeColumn.Items.IndexOf(column.DataType.ToString().ToUpper())
            If dataTypeIndex <> -1 Then
                cell.Value = dataTypeColumn.Items(dataTypeIndex)
            End If
        Next
    End Sub
    Private Function DoesTableExist(tableName As String, connection As SqlConnection) As Boolean
        Dim commandText As String = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName"
        Using command As New SqlCommand(commandText, connection)
            command.Parameters.AddWithValue("@TableName", tableName)
            Return Convert.ToInt32(command.ExecuteScalar()) > 0
        End Using
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub



    Private Sub RefreshData()
        Dim numberofrows As String = ""
        If Not String.IsNullOrWhiteSpace(txtNumberofRow.Text) Then
            numberofrows = " TOP " & txtNumberofRow.Text
        End If
        Try
            If cmbTables.SelectedItem IsNot Nothing Then
                Dim selectedTable As String = cmbTables.SelectedItem.ToString()
                Dim whereClause As String = ""
                If txtWhereConditiontext.Text.Trim() = "" Then
                    whereClause = ""
                ElseIf cmbFilterColumnsnames.SelectedItem IsNot Nothing AndAlso ComboOperators.SelectedItem IsNot Nothing Then
                    Dim selectedColumn As String = cmbFilterColumnsnames.SelectedItem.ToString()
                    Dim selectedOperator As String = ComboOperators.SelectedItem.ToString()
                    Dim value As String = txtWhereConditiontext.Text
                    If Not IsNumeric(value) Then
                        value = "'" & value & "'"
                    End If
                    whereClause = $" WHERE {selectedColumn} {selectedOperator} {value}"
                End If
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT " & numberofrows & " * FROM " & selectedTable & whereClause
                    txtShowmecode.Text = query
                    Dim adapter As New SqlDataAdapter(query, connection)
                    Dim dataSet As New DataSet()
                    adapter.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                End Using
            Else
                ErrorPage.Show()
                customMsgBoxF.txtMsgSucess.Text = "Please select a table first."
            End If
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = ex.Message
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCustumCodeSQL.Click
        If refreshActiveIndicator = "1" Then
            Timer1.Start()
            countdown = RefreshTime
        End If

        Dim numberofrows As String = ""
        Dim parameterName As String = ""
        Dim parameterValue As Object = Nothing

        Try


            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = txtShowmecode.Text


                Using command As New SqlCommand(query, connection)


                    Dim adapter As New SqlDataAdapter(command)
                    Dim dataSet As New DataSet()
                    adapter.Fill(dataSet)

                    If dataSet.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("The query returned no results.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        DataGridView1.DataSource = dataSet.Tables(0)
                    End If
                End Using
            End Using

        Catch ex As SqlException
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = "Database Error: " & ex.Message
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = "An error occurred: " & ex.Message
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            ' Create PrintDocument and set landscape orientation
            Dim printDocument As New PrintDocument()
            printDocument.DefaultPageSettings.Landscape = True

            ' Initialize fonts
            Dim basicFont As New Font("Arial", 10)
            Dim headerFont As New Font("Arial", 12, FontStyle.Bold)

            ' Calculate fixed column widths based on header text (avoid wrapping headers)
            Dim columnWidths As New List(Of Integer)
            For Each column As DataGridViewColumn In DataGridView1.Columns
                Dim headerWidth As Integer = CInt(ePrint.Graphics.MeasureString(column.HeaderText, headerFont).Width) + 10 ' Add padding
                columnWidths.Add(headerWidth)
            Next

            ' Adjust column widths based on cell content
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                        Dim cellValue As String = If(row.Cells(i).Value IsNot Nothing, row.Cells(i).Value.ToString(), "")
                        Dim cellWidth As Integer = CInt(ePrint.Graphics.MeasureString(cellValue, basicFont).Width) + 10 ' Add padding
                        If cellWidth > columnWidths(i) Then
                            columnWidths(i) = cellWidth
                        End If
                    Next
                End If
            Next

            ' Set the print document handler
            AddHandler printDocument.PrintPage, Sub(senderPrint As Object, ePrint As PrintPageEventArgs)
                                                    Dim yPos As Integer = 50        ' Start printing below the top margin
                                                    Dim startX As Integer = 50       ' Start printing from the left margin
                                                    Dim cellPadding As Integer = 5
                                                    Dim rowsPerPage As Integer = CalculateRowsPerPage(ePrint, basicFont)

                                                    ' Print the header
                                                    For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                                        Dim alignment As StringFormat = New StringFormat()
                                                        alignment.Alignment = If(DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight, StringAlignment.Far, StringAlignment.Near)
                                                        ePrint.Graphics.DrawString(DataGridView1.Columns(i).HeaderText, headerFont, Brushes.Black, New RectangleF(startX, yPos, columnWidths(i), headerFont.Height), alignment)
                                                        startX += columnWidths(i) + cellPadding
                                                    Next
                                                    yPos += headerFont.Height + cellPadding


                                                    ' Print each row
                                                    Dim rowIndex As Integer = 0
                                                    While rowIndex < DataGridView1.Rows.Count AndAlso rowIndex < rowsPerPage
                                                        startX = 50  ' Reset for each row
                                                        Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
                                                        If Not row.IsNewRow Then
                                                            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                                                                Dim cellValue As String = If(row.Cells(i).Value IsNot Nothing, row.Cells(i).Value.ToString(), "")

                                                                Dim alignment As StringFormat = New StringFormat()
                                                                alignment.Alignment = If(DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight, StringAlignment.Far, StringAlignment.Near)
                                                                ePrint.Graphics.DrawString(cellValue, basicFont, Brushes.Black, New RectangleF(startX, yPos, columnWidths(i), basicFont.Height), alignment)
                                                                startX += columnWidths(i) + cellPadding
                                                            Next
                                                            yPos += basicFont.Height + cellPadding
                                                        End If
                                                        rowIndex += 1

                                                        ' Check for page break
                                                        If yPos + basicFont.Height > ePrint.MarginBounds.Bottom Then
                                                            ePrint.HasMorePages = True
                                                            Exit Sub
                                                        End If
                                                    End While
                                                End Sub

            ' ... (Rest of the code for PrintPreviewDialog) ...
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing: " & ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Helper Function: Calculate rows per page
    Private Function CalculateRowsPerPage(e As PrintPageEventArgs, font As Font) As Integer
        Return CInt(e.MarginBounds.Height / (font.Height + 5)) ' 5 is for padding
    End Function




End Class
