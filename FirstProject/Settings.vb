Imports System.Data.SqlClient
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class Settings
    Private Sub LoadConnectionSettings()
        Dim configFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "settings.xml")
        If File.Exists(configFile) Then
            Try
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
            Catch ex As Exception
                Throw New Exception("Error occurred while loading connection settings", ex)
            End Try
        Else
            ' If the file doesn't exist, create a default one
            CreateDefaultConfigFile(configFile)
        End If
    End Sub



    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        Try
            PopulateDatabaseList()

        Catch ex As SqlException
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = $"SQL Connection Error: {ex.Message}"

        Catch ex As Exception
            ErrorPage.Show()
            ErrorPage.txtErrorMassage.Text = $"General Connection Error: {ex.Message}"
        End Try
    End Sub

    Private Sub PopulateDatabaseList()
        Dim connectionString As String = $"Data Source={txtdatabaseName.Text};Initial Catalog=master;User ID={txtUsernameSQL.Text};Password={txtPassword.Text};"

        Using connection As New SqlConnection(connectionString)
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

            Catch ex As SqlException
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = $"SQL Error (database listing): {ex.Message}"
            Catch ex As Exception
                ErrorPage.Show()
                ErrorPage.txtErrorMassage.Text = $"Error (database listing): {ex.Message}"
            Finally
                connection.Close()
            End Try
        End Using
    End Sub
    Private Sub btnSavesettings_Click(sender As Object, e As EventArgs) Handles btnSavesettings.Click
        Try
            Dim configFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "settings.xml")
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
            ' Show message and restart
            MessageBox.Show("Connection settings saved successfully! The application will now restart to apply the changes.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Application.Restart()

        Catch ex As Exception
            MessageBox.Show($"Error occurred while saving connection settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNewConnectionSettings_Click(sender As Object, e As EventArgs) Handles btnNewConnectionSettings.Click
        txtdatabaseName.Text = ""
        txtUsernameSQL.Text = ""
        txtPassword.Text = ""
        ComboBox1.SelectedIndex = -1 ' Clear selected database in ComboBox
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUsersTable.Click
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

    Private Function HashPassword(password As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim hashBytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(password))
            ' Convert to a suitable format for storing (e.g., Base64)
            Return Convert.ToBase64String(hashBytes)
        End Using
    End Function
    Private Sub AddSetting(doc As XmlDocument, parent As XmlNode, key As String, value As String)
        Dim settingNode As XmlNode = doc.CreateElement("add")
        Dim keyAttr As XmlAttribute = doc.CreateAttribute("key")
        keyAttr.Value = key
        Dim valueAttr As XmlAttribute = doc.CreateAttribute("value")
        valueAttr.Value = value
        settingNode.Attributes.Append(keyAttr)
        settingNode.Attributes.Append(valueAttr)
        parent.AppendChild(settingNode)
    End Sub


    Private Sub CreateDefaultConfigFile(configFile As String)
        ' Ensure the directory exists
        Dim configDirectory As String = Path.GetDirectoryName(configFile)
        If Not Directory.Exists(configDirectory) Then
            Directory.CreateDirectory(configDirectory)
        End If

        ' Create the basic XML structure
        Dim xmlDoc As New XmlDocument()

        Dim declaration As XmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", Nothing)
        xmlDoc.AppendChild(declaration)

        Dim rootNode As XmlNode = xmlDoc.CreateElement("configuration")
        xmlDoc.AppendChild(rootNode)

        Dim appSettingsNode As XmlNode = xmlDoc.CreateElement("appSettings")
        rootNode.AppendChild(appSettingsNode)

        ' Add default settings nodes
        AddSetting(xmlDoc, appSettingsNode, "DataSource", " ")
        AddSetting(xmlDoc, appSettingsNode, "Username", " ")
        AddSetting(xmlDoc, appSettingsNode, "Password", " ")
        AddSetting(xmlDoc, appSettingsNode, "InitialCatalog", " ")

        ' Save the XML document
        xmlDoc.Save(configFile)
    End Sub

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim configFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "settings.xml")
        If File.Exists(configFile) Then
            Try
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
            Catch ex As Exception
                Throw New Exception("Error occurred while loading connection settings", ex)
            End Try
        Else
            ' If the file doesn't exist, create a default one
            CreateDefaultConfigFile(configFile)
        End If
    End Sub

    Private Sub ToolStripStatusLabel2_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel2.Click

    End Sub
End Class
