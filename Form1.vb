Imports MySql.Data.MySqlClient
<<<<<<<<< Temporary merge branch 1

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
=========
Imports System.IO
Imports System.Drawing

Public Class Form1
    ' Connection string - shared
    Private connStr As String = "Server=192.168.0.122; Port=3306; Username=root; Password=astronext@2025; Database=jonbalsatdb;"

    ' Button1: Fetch sensor data
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using conn As New MySqlConnection(connStr)
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT SN, Temperature, Humidity, Pressure, created_at FROM jonbalsattb WHERE SN=@parm1", conn)
                cmd.Parameters.AddWithValue("@parm1", TextBox1.Text)

                Dim reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    TextBox2.Text = If(IsDBNull(reader("SN")), "", reader("SN").ToString())
                    TextBox8.Text = If(IsDBNull(reader("Temperature")), "", reader("Temperature").ToString())
                    TextBox3.Text = If(IsDBNull(reader("Humidity")), "", reader("Humidity").ToString())
                    TextBox4.Text = If(IsDBNull(reader("Pressure")), "", reader("Pressure").ToString())
                    TextBox5.Text = If(IsDBNull(reader("created_at")), "", reader("created_at").ToString())
                Else
                    MessageBox.Show("No sensor data found.")
                End If
                reader.Close()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub

    ' Button2: Fetch image from pi_images
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Using conn As New MySqlConnection(connStr)
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT image, filename FROM pi_images WHERE id=@imgId", conn)
                cmd.Parameters.AddWithValue("@imgId", TextBox6.Text)

                Dim reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.Read() AndAlso Not IsDBNull(reader("image")) Then
                    TextBox7.Text = If(IsDBNull(reader("filename")), "", reader("filename").ToString())
                    Dim imgData As Byte() = CType(reader("image"), Byte())
                    Using ms As New MemoryStream(imgData)
                        PictureBox1.Image = Image.FromStream(ms)
                    End Using
                Else
                    PictureBox1.Image = Nothing
                    MessageBox.Show("No image found.")
                End If
                reader.Close()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
>>>>>>>>> Temporary merge branch 2
    End Sub
End Class
