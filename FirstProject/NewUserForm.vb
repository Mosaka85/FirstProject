Imports System.Data.SqlClient
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class NewUserForm

    Private Sub btnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
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

                ' Check for empty fields and password match
                If txtnewusername.Text = "" Then
                    ShowErrorMessage("Please Enter Username")
                ElseIf txtNewpassword1.Text = "" OrElse txtnewpassword2.Text = "" Then
                    ShowErrorMessage("Please Enter Valid Passwords")
                ElseIf txtNewpassword1.Text <> txtnewpassword2.Text Then
                    ShowErrorMessage("Passwords Do Not Match")
                    txtNewpassword1.Text = ""
                    txtnewpassword2.Text = ""
                Else
                    ' Check if the username already exists
                    Dim userExistsQuery As String = "SELECT COUNT(*) FROM USERS_LOGINS WHERE Username = @Username"
                    Using cmdCheckUser As New SqlCommand(userExistsQuery, Con)
                        cmdCheckUser.Parameters.AddWithValue("@Username", txtnewusername.Text)
                        Dim userCount As Integer = Convert.ToInt32(cmdCheckUser.ExecuteScalar())
                        If userCount > 0 Then
                            ShowErrorMessage("Username Already Exists")
                        Else
                            ' Hash the password
                            Dim passwordHash As String = HashPassword(txtNewpassword1.Text)

                            ' Insert new user into the database
                            Dim query As String = "INSERT INTO USERS_LOGINS (Username, Password) VALUES (@Username, @Password)"
                            Using cmd As New SqlCommand(query, Con)
                                cmd.Parameters.AddWithValue("@Username", txtnewusername.Text)
                                cmd.Parameters.AddWithValue("@Password", passwordHash)

                                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                                If rowsAffected > 0 Then
                                    MessageBox.Show("User registered successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                    Me.Hide()
                                Else
                                    MessageBox.Show("Failed to register user.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End If
                                txtnewusername.Text = ""
                                txtNewpassword1.Text = ""
                                txtnewpassword2.Text = ""
                            End Using
                        End If
                    End Using
                End If
            End Using
        Catch ex As Exception
            ShowErrorMessage("Error: " & ex.Message)
        End Try
    End Sub
    Private Sub ShowErrorMessage(message As String)
        ErrorPage.Show()
        ErrorPage.txtErrorMassage.Text = message
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Hide()
    End Sub


    Private Function HashPassword(password As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim hashBytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(password))
            ' Convert to a suitable format for storing (e.g., Base64)
            Return Convert.ToBase64String(hashBytes)
        End Using
    End Function

    Private Sub NewUserForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtnewusername.Text = ""
        txtNewpassword1.Text = ""
        txtnewpassword2.Text = ""
    End Sub
End Class


