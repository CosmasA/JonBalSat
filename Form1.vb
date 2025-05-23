Imports MySql.Data.MySqlClient

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn As New MySqlConnection("Server=192.168.0.122; Port=3306; Username=root; Password=astronext@2025; Database=jonbalsatdb;")
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("SELECT SN, Temperature, Humidity, Pressure, created_at FROM jonbalsattb WHERE SN=@parm1", conn)
            cmd.Parameters.AddWithValue("@parm1", TextBox1.Text)

            Dim myreader As MySqlDataReader = cmd.ExecuteReader()
            If myreader.Read() Then
                TextBox2.Text = If(IsDBNull(myreader("SN")), "", myreader("SN").ToString())
                TextBox3.Text = If(IsDBNull(myreader("Temperature")), "", myreader("Temperature").ToString())
                TextBox4.Text = If(IsDBNull(myreader("Humidity")), "", myreader("Humidity").ToString())
                TextBox5.Text = If(IsDBNull(myreader("created_at")), "", myreader("created_at").ToString())
            Else
                MessageBox.Show("No Data Found")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
End Class
