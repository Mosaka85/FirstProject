Imports System.Data.SqlClient
Imports System.Xml

Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ToolStripStatusLabel1.Text = "Version: V1.00 Build 00.00" & My.Application.Info.Version.ToString()
        ToolStripStatusLabel2.Text = "© 2024 MOSAKA PTY LTD. All rights reserved."
        Try
            ' Load connection settings from XML
            LoadConnectionSettings()
        Catch ex As Exception
            MessageBox.Show($"Error occurred while loading connection settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub LoadConnectionSettings()
        Dim configFile As String = "C:\Users\TSHEP\source\repos\FirstProject\FirstProject\settings.xml" ' Assuming the XML file is in the project directory
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(configFile)
        Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='DataSource']")
        Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Username']")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Password']")
        Dim databaseNameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='InitialCatalog']")
        txtdatabaseName.Text = dataSourceNode.Attributes("value").Value
        txtUsernameSQL.Text = usernameNode.Attributes("value").Value
        txtPassword.Text = passwordNode.Attributes("value").Value
        Dim savedDatabaseName As String = databaseNameNode.Attributes("value").Value
        If ComboBox1.Items.Contains(savedDatabaseName) Then
            ComboBox1.SelectedItem = savedDatabaseName
        End If
    End Sub

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        Try
            PopulateDatabaseList()
        Catch ex As Exception
            MessageBox.Show($"Error occurred while loading databases: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub PopulateDatabaseList()
        Dim connectionString As String = $"Data Source={txtdatabaseName.Text};Initial Catalog=FirstProject;User ID={txtUsernameSQL.Text};Password={txtPassword.Text};"
        Dim connection As New SqlConnection(connectionString)

        Try
            connection.Open()
            Dim databases As New List(Of String)
            Dim command As New SqlCommand("SELECT name FROM sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')", connection)
            Using reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    databases.Add(reader.GetString(0))
                End While
            End Using
            ComboBox1.DataSource = databases
            customMsgBoxF.Show()
            customMsgBoxF.txtMsgSucess.Text = "Connection successful!" & "Success"

        Catch ex As Exception
            ErrorPage.Show()
            Dim Errorcatch As String = $"Connection failed!{Environment.NewLine}{ex.Message}"
            ErrorPage.txtErrorMassage.Text = Errorcatch
        Finally
            connection.Close()
        End Try
    End Sub
    Private Sub btnSavesettings_Click(sender As Object, e As EventArgs) Handles btnSavesettings.Click
        Try
            Dim configFile As String = "C:\Users\TSHEP\source\repos\FirstProject\FirstProject\settings.xml" ' Assuming the XML file is in the project directory
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(configFile)
            Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='DataSource']")
            dataSourceNode.Attributes("value").Value = txtdatabaseName.Text
            Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Username']")
            usernameNode.Attributes("value").Value = txtUsernameSQL.Text
            Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Password']")
            passwordNode.Attributes("value").Value = txtPassword.Text
            Dim databaseNameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='InitialCatalog']")
            databaseNameNode.Attributes("value").Value = ComboBox1.SelectedItem.ToString()
            xmlDoc.Save(configFile)
            MessageBox.Show("Connection settings saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Error occurred while saving connection settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Dim tableName As String = "USERS_LOGINS"
        Dim connectionString As String = $"Data Source={txtdatabaseName.Text};Initial Catalog={ComboBox1.SelectedItem.ToString()};User ID={txtUsernameSQL.Text};Password={txtPassword.Text};"
        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim checkTableQuery As String = $"IF OBJECT_ID('{tableName}', 'U') IS NULL
                                                    SELECT 0
                                                ELSE
                                                    SELECT 1"
            Using command As New SqlCommand(checkTableQuery, connection)
                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing AndAlso Convert.ToInt32(result) = 0 Then
                    ' Table doesn't exist, create it
                    Dim createTableQuery As String = $"
                        CREATE TABLE [dbo].[USERS_LOGINS](
	                        [UserName] [varchar](50) NULL,
	                        [PASSWORD] [varchar](50) NULL,
	                        [ID] [int] IDENTITY(1,1) NOT NULL,
                         CONSTRAINT [PK_USERS_LOGINS] PRIMARY KEY CLUSTERED 
                        (
	                        [ID] ASC
                        )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
                        ) ON [PRIMARY]
                        "
                    Using createCommand As New SqlCommand(createTableQuery, connection)
                        createCommand.ExecuteNonQuery()
                        MessageBox.Show("Table created successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Using
                Else
                    MessageBox.Show("Table already exists.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using
    End Sub
    Private Sub btnNewConnectionSettings_Click(sender As Object, e As EventArgs) Handles btnNewConnectionSettings.Click
        txtdatabaseName.Text = ""
        txtUsernameSQL.Text = ""
        txtPassword.Text = ""
        ComboBox1.SelectedIndex = -1 ' Clear selected database in ComboBox
    End Sub
End Class
