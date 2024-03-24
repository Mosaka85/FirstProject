Imports System.Data.SqlClient
Imports System.Xml
Imports FirstProject.Excel
Imports System.Linq
Imports System.Data.OleDb
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Xml.XPath
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports DataTable = System.Data.DataTable
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Workbook = Microsoft.Office.Interop.Excel.Workbook
Imports Application = Microsoft.Office.Interop.Excel.Application
Imports Microsoft.SqlServer.Management.Smo
Imports Microsoft.SqlServer.Management.Common

Public Class menuL
    Dim configFile As String = "C:\Users\TSHEP\source\repos\FirstProject\FirstProject\settings.xml"
    Dim xmlDoc As New XmlDocument()
    Dim connectionString As String
    Private bgWorker As Object
    Private excelFilePath As String
    Private Sub menuL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        xmlDoc.Load(configFile)
        ProgressBar1.Visible = False
        btnImport.Visible = False
        txtSQLtablename.Visible = False
        txtNewSQLTable.Visible = False
        lblSELECTWORKSHEET.Visible = False
        cmbSELECTsheet.Visible = False
        txtdatabasename.Visible = False
        btnCreatedabase.Visible = False
        lblCreatedatabase.Visible = False
        Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='DataSource']")
        Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Username']")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Password']")
        Dim databaseNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='InitialCatalog']")
        Dim dataSource As String = dataSourceNode.Attributes("value").Value
        Dim username As String = usernameNode.Attributes("value").Value
        Dim password As String = passwordNode.Attributes("value").Value
        Dim database As String = databaseNode.Attributes("value").Value
        txtDabasenameSHOW.Text = database
        connectionString = $"Data Source={dataSource};Initial Catalog={database};User ID={username};Password={password};"
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
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DataGridView1.DataSource = Nothing
        If cmbTables.SelectedItem IsNot Nothing Then
            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim selectedTable As String = cmbTables.SelectedItem.ToString()
                    Dim query As String = $"SELECT * FROM {selectedTable}"
                    Dim adapter As New SqlDataAdapter(query, connection)
                    Dim dataSet As New DataSet()
                    adapter.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                End Using
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
    End Sub
    Private Sub btnShowtable_Click(sender As Object, e As EventArgs) Handles btnShowtable.Click

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
                customMsgBoxF.Show()
                customMsgBoxF.txtMsgSucess.Text = "Please select a table first."
            End If
        Catch ex As Exception
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = ex.Message
        End Try
    End Sub



    Private oExcel As Microsoft.Office.Interop.Excel.Application
    Private oBook As Microsoft.Office.Interop.Excel.Workbook
    Private oSheet As Microsoft.Office.Interop.Excel.Worksheet

    Private Sub BtnExportData_Click(sender As Object, e As EventArgs) Handles btnExportData.Click
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
                Dim schemaName As String = If(showViews, "Views", "Tables")
                Dim tablesSchema As DataTable = If(showViews, connection.GetSchema("Views"), connection.GetSchema("Tables"))
                If tablesSchema IsNot Nothing AndAlso tablesSchema.Rows.Count > 0 Then
                    For Each row As DataRow In tablesSchema.Rows
                        Dim tableName As String = If(showViews, row("TABLE_NAME").ToString(), row("TABLE_NAME").ToString())
                        cmbTables.Items.Add(tableName)
                    Next
                Else
                    MessageBox.Show($"No {(If(showViews, "views", "tables"))} found in the database.")
                End If
            End Using
        Catch ex As Exception
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = "Error: " & ex.Message
        End Try
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
        txtNumberofRow.Text = ""
        If cmbTables.SelectedItem IsNot Nothing Then
            Dim selectedTable As String = cmbTables.SelectedItem.ToString()
            PopulateColumnsComboBox(selectedTable)
        End If

        FillOperatorsCombo(ComboOperators)
    End Sub

    Private Sub btnCreatletable_Click(sender As Object, e As EventArgs) Handles btnCreatletable.Click
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
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If cmbSELECTsheet.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(txtNewSQLTable.Text) Then
            MessageBox.Show("Please select a sheet and specify a SQL table name before importing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim selectedSheet As String = cmbSELECTsheet.SelectedItem.ToString()
        Dim excelConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"

        Using excelConnection As New OleDb.OleDbConnection(excelConnectionString)
            excelConnection.Open()
            Dim dt As New DataTable()
            Dim query As String = "SELECT * FROM [" & selectedSheet & "$]"
            Using adapter As New OleDbDataAdapter(query, excelConnection)
                adapter.Fill(dt)
            End Using
            Using sqlConnection As New SqlConnection(connectionString)
                sqlConnection.Open()
                Dim tableName As String = txtSQLtablename.Text
                Dim createTableQuery As String = "CREATE TABLE " & tableName & " ("
                For Each column As DataColumn In dt.Columns
                    createTableQuery &= "[" & column.ColumnName & "] NVARCHAR(MAX), "
                Next
                createTableQuery = createTableQuery.TrimEnd(", ") & ")"
                txtShowmecode.Text = createTableQuery
                Using createTableCommand As New SqlCommand(createTableQuery, sqlConnection)
                    createTableCommand.ExecuteNonQuery()
                End Using
                Using bulkCopy As New SqlBulkCopy(sqlConnection)
                    bulkCopy.DestinationTableName = tableName
                    For Each column As DataColumn In dt.Columns
                        bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName)
                    Next
                    bulkCopy.WriteToServer(dt)
                End Using
                MessageBox.Show("Data imported successfully to SQL Server table: " & tableName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
            btnImport.Visible = False
            txtSQLtablename.Visible = False
            txtNewSQLTable.Visible = False
            lblSELECTWORKSHEET.Visible = False
            cmbSELECTsheet.Visible = False
        End Using
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
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Hide()
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
End Class