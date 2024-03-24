Imports System.Data.SqlClient
Imports System.Xml

Public Class LOGINPAGE
    Private Sub btnLogIn_Click(sender As Object, e As EventArgs) Handles btnLogIn.Click
        If txtUserName.Text = "" Then
            ShowErrorMessage("Please Enter Username")
            Exit Sub
        End If
        If txtPassword.Text = "" Then
            ShowErrorMessage("Invalid Password")
            Exit Sub
        End If
        Try
            Dim configFile As String = "C:\Users\TSHEP\source\repos\FirstProject\FirstProject\settings.xml"
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
                Dim rdr As SqlDataReader
                Dim cmd As New SqlCommand("SELECT * FROM USERS_LOGINS WHERE Username=@Username AND Password=@Password", Con)
                cmd.Parameters.AddWithValue("@Username", txtUserName.Text)
                cmd.Parameters.AddWithValue("@Password", txtPassword.Text)
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    menuL.Show()
                    Me.Hide()
                    txtUserName.Text = ""
                    txtPassword.Text = ""
                Else
                    ShowErrorMessage("Invalid Correct Username Or password")
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
End Class