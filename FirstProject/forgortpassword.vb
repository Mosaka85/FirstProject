Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Data.SqlClient ' Add this import for SqlCommand
Imports System.IO

Public Class forgortpassword


    Private Sub btnupdatepassword_Click(sender As Object, e As EventArgs) Handles btnupdatepassword.Click
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

        Dim newPassword = txtNewPassword.Text
        Dim newpasswordConfirm = Txtnewpassword1.Text ' Assuming this is the confirmation field

        ' Validate password and confirmation match
        If newPassword <> newpasswordConfirm Then
            MessageBox.Show("Passwords do not match.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ' Exit the function if there's no match
        End If

        ' Validate that the new password is not empty 
        If newPassword.Length = 0 Then
            MessageBox.Show("Please enter a new password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim newPasswordHash As String = HashPassword(newPassword)

        ' Assuming 'Con' is your database connection
        Using Con As New SqlConnection(connectionString)
            Con.Open()
            ' Change the query to an UPDATE
            Dim query As String = "UPDATE USERS_LOGINS SET Password = @NewPassword WHERE Username = @Username"
            Using cmd As New SqlCommand(query, Con)
                cmd.Parameters.AddWithValue("@Username", txtUsernameChange.Text)
                cmd.Parameters.AddWithValue("@NewPassword", newPasswordHash)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                If rowsAffected > 0 Then
                    MessageBox.Show("Password changed successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Failed to change password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        End Using
    End Sub

    Private Function HashPassword(password As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim hashBytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(password))
            Return Convert.ToBase64String(hashBytes)
        End Using
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnbackChangepassword.Click
        Me.Hide()
        LOGINPAGE.Show()
    End Sub

    Private Sub Txtnewpassword1_TextChanged(sender As Object, e As EventArgs) Handles Txtnewpassword1.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub


End Class
