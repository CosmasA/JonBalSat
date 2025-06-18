Imports System.Drawing
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Form2
    ' MySQL connection string
    Private connStr As String = "Server=192.168.0.122; Port=3306; Username=root; Password=astronext@2025; Database=jonbalsatdb;"

    ' Load all images and details into FlowLayoutPanel
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        flowImages.Controls.Clear() ' Clear previous items

        Using conn As New MySqlConnection(connStr)
            Try
                conn.Open()
                Dim query As String = "SELECT id, image, filename, created_at FROM pi_images ORDER BY created_at DESC LIMIT 20"
                Dim cmd As New MySqlCommand(query, conn)
                Dim reader As MySqlDataReader = cmd.ExecuteReader()

                While reader.Read()
                    Dim imgData As Byte() = CType(reader("image"), Byte())
                    Dim filename As String = reader("filename").ToString()
                    Dim createdAt As String = reader("created_at").ToString()

                    ' Convert byte array to image
                    Dim img As Image = Nothing
                    Using ms As New MemoryStream(imgData)
                        img = Image.FromStream(ms)
                    End Using

                    ' Create PictureBox
                    Dim pb As New PictureBox With {
                        .Image = img,
                        .Width = 150,
                        .Height = 100,
                        .SizeMode = PictureBoxSizeMode.Zoom,
                        .Margin = New Padding(5)
                    }

                    ' Create Label for filename and timestamp
                    Dim lbl As New Label With {
                        .Text = $"{filename}{vbCrLf}{createdAt}",
                        .AutoSize = True
                    }

                    ' Container panel to hold image and text
                    Dim container As New Panel With {
                        .Width = 160,
                        .Height = 150
                    }
                    container.Controls.Add(pb)
                    container.Controls.Add(lbl)
                    lbl.Top = pb.Bottom + 5

                    ' Add to FlowLayoutPanel
                    flowImages.Controls.Add(container)
                End While

                reader.Close()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
End Class
