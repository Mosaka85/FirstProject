Imports System.Data.SqlClient
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml

Public Class LOGINPAGE
    Private Sub btnLogIn_Click(sender As Object, e As EventArgs) Handles btnLogIn.Click
        If txtUserName.Text = "" Then
            ShowErrorMessage("Please Enter Username")
            Exit Sub
        End If
        MainMenu.btnCreatedatabase.Visible = False
        If txtUserName.Text = "Mosaka" Then
            MainMenu.btnCreatedatabase.Visible = True

        End If
        If txtPassword.Text = "" Then
            ShowErrorMessage("Invalid Password")
            Exit Sub
        End If
        Try
            Dim configFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "settings.xml")
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(configFile)
            Dim dataSourceNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='DataSource']")
            Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Username']")
            Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='Password']")
            Dim databaseNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='InitialCatalog']")
            Dim dataSource As String = dataSourceNode.Attributes("value").Value
            Dim username As String = usernameNode.Attributes("value").Value
            Dim password As String = passwordNode.Attributes("value").Value
            Dim database As String = databaseNode.Attributes("value").Value

            Dim connectionString As String = $"Data Source={dataSource};Initial Catalog={database};User ID={username};Password={password};"
            Using Con As New SqlConnection(connectionString)
                Con.Open()
                ' Hash the password provided by the user
                Dim hashedPassword As String = HashPassword(txtPassword.Text)

                ' Check if the user exists
                Dim query As String = "SELECT COUNT(*) FROM USERS_LOGINS WHERE Username=@Username AND Password=@Password"
                Using cmd As New SqlCommand(query, Con)
                    cmd.Parameters.AddWithValue("@Username", txtUserName.Text)
                    cmd.Parameters.AddWithValue("@Password", hashedPassword)

                    ' Execute the query to check if the user exists
                    Dim userCount As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    If userCount > 0 Then
                        ' User exists, show the menu
                        MainMenu.Show()
                        Me.Hide()
                        txtUserName.Text = ""
                        txtPassword.Text = ""
                    Else
                        ' User does not exist
                        ShowErrorMessage("Incorrect Username or Password")
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowErrorMessage("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowErrorMessage(message As String)
        ErrorPage.Show()
        ErrorPage.txtErrorMassage.Text = message
    End Sub

    Private Sub btnEXIT_Click(sender As Object, e As EventArgs) Handles btnEXIT.Click
        Application.Exit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NewUserForm.Show()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Settings.Show()
    End Sub

    Private Function HashPassword(password As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim hashBytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(password))
            ' Convert to a suitable format for storing (e.g., Base64)
            Return Convert.ToBase64String(hashBytes)
        End Using
    End Function

    Private Sub btnChangePassword_Click(sender As Object, e As EventArgs) Handles btnChangePassword.Click
        Me.Hide()
        forgortpassword.Show()
    End Sub
End Class