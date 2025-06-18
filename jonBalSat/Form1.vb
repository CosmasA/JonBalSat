Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.IO
Imports System.Text
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class Form1
    ' Connection string - shared
    Private connStr As String = "Server=192.168.0.122;Port=3306;Uid=root;Pwd=astronext@2025;Database=jonbalsatdb;"

    ' Form Load Event: Fetch sensor data
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSensorData()
        LoadMPUData()
    End Sub

    Private Sub LoadSensorData()
        Using conn As New MySqlConnection(connStr)
            Try
                conn.Open()
                Dim query As String = "SELECT SN, Temperature, Humidity, Pressure, created_at FROM jonbalsattb ORDER BY SN DESC LIMIT 200"
                Dim cmd As New MySqlCommand(query, conn)
                Dim adapter As New MySqlDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)

                If table.Rows.Count > 0 Then
                    DataGridView1.DataSource = table
                    FormatDataGridView()
                Else
                    MessageBox.Show("No sensor data found.")
                    DataGridView1.DataSource = Nothing
                End If

            Catch ex As MySqlException
                MessageBox.Show("MySQL Error: " & ex.Message & vbCrLf & "Error Code: " & ex.Number)
            Catch ex As Exception
                MessageBox.Show("General Error: " & ex.Message)
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End Using
    End Sub


    Private Sub LoadMPUData()
        Using conn As New MySqlConnection(connStr)
            Try
                conn.Open()
                Dim query As String = "SELECT ID, Accelerometer, Gyroscope, Created FROM mputable ORDER BY ID DESC LIMIT 200"
                Dim cmd As New MySqlCommand(query, conn)
                Dim adapter As New MySqlDataAdapter(cmd)
                Dim table As New DataTable()
                adapter.Fill(table)

                If table.Rows.Count > 0 Then
                    DataGridView2.DataSource = table
                    FormatDataGridView2()
                Else
                    MessageBox.Show("No sensor data found.")
                    DataGridView2.DataSource = Nothing
                End If

            Catch ex As MySqlException
                MessageBox.Show("MySQL Error: " & ex.Message & vbCrLf & "Error Code: " & ex.Number)
            Catch ex As Exception
                MessageBox.Show("General Error: " & ex.Message)
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End Using
    End Sub

    Private Sub FormatDataGridView()
        Try
            With DataGridView1
                If .Columns.Contains("created_at") Then
                    .Columns("created_at").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
                    .Columns("created_at").HeaderText = "Created At"
                End If

                If .Columns.Contains("SN") Then
                    .Columns("SN").HeaderText = "S/No"
                    .Columns("SN").Width = 100
                End If

                If .Columns.Contains("Temperature") Then
                    .Columns("Temperature").DefaultCellStyle.Format = "N2"
                    .Columns("Temperature").HeaderText = "Temperature (°C)"
                End If

                If .Columns.Contains("Humidity") Then
                    .Columns("Humidity").DefaultCellStyle.Format = "N2"
                    .Columns("Humidity").HeaderText = "Humidity (%)"
                End If

                If .Columns.Contains("Pressure") Then
                    .Columns("Pressure").DefaultCellStyle.Format = "N2"
                    .Columns("Pressure").HeaderText = "Pressure"
                End If

                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .ReadOnly = True
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray
            End With
        Catch ex As Exception
            MessageBox.Show("Error formatting DataGridView: " & ex.Message)
        End Try
    End Sub

    Private Sub FormatDataGridView2()
        Try
            With DataGridView2
                If .Columns.Contains("Created") Then
                    .Columns("Created").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
                    .Columns("Created").HeaderText = "Created At"
                End If

                If .Columns.Contains("ID") Then
                    .Columns("ID").HeaderText = "ID"
                    .Columns("ID").Width = 100
                End If

                If .Columns.Contains("Accelerometer") Then
                    .Columns("Accelerometer").DefaultCellStyle.Format = "N2"
                    .Columns("Accelerometer").HeaderText = "Accelerometer"
                End If

                If .Columns.Contains("Gyroscope") Then
                    .Columns("Gyroscope").DefaultCellStyle.Format = "N2"
                    .Columns("Gyroscope").HeaderText = "Gyroscope"
                End If

                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .ReadOnly = True
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray
            End With
        Catch ex As Exception
            MessageBox.Show("Error formatting DataGridView: " & ex.Message)
        End Try
    End Sub

    ' Refresh buttons
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadSensorData()
    End Sub
    Private Sub refreshBtn_Click(sender As Object, e As EventArgs) Handles refreshBtn.Click
        LoadMPUData()
    End Sub


    ' Export button with format selection
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If DataGridView1.DataSource Is Nothing OrElse DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("No data to export!")
            Return
        End If

        Dim exportDialog As New Form()
        exportDialog.Text = "Select Export Format"
        exportDialog.Size = New Size(300, 200)
        exportDialog.StartPosition = FormStartPosition.CenterParent

        Dim btnCSV As New Button() With {.Text = "Export to CSV", .Location = New Point(20, 20), .Size = New Size(120, 30)}
        Dim btnExcel As New Button() With {.Text = "Export to Excel", .Location = New Point(150, 20), .Size = New Size(120, 30)}
        Dim btnWord As New Button() With {.Text = "Export to Word", .Location = New Point(20, 60), .Size = New Size(120, 30)}
        Dim btnPDF As New Button() With {.Text = "Export to PDF", .Location = New Point(150, 60), .Size = New Size(120, 30)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .Location = New Point(85, 100), .Size = New Size(120, 30)}

        AddHandler btnCSV.Click, Sub()
                                     ExportToCSV()
                                     exportDialog.Close()
                                 End Sub
        AddHandler btnExcel.Click, Sub()
                                       ExportToExcel()
                                       exportDialog.Close()
                                   End Sub
        AddHandler btnWord.Click, Sub()
                                      ExportToWord()
                                      exportDialog.Close()
                                  End Sub
        AddHandler btnPDF.Click, Sub()
                                     ExportToPDF()
                                     exportDialog.Close()
                                 End Sub
        AddHandler btnCancel.Click, Sub() exportDialog.Close()

        exportDialog.Controls.AddRange({btnCSV, btnExcel, btnWord, btnPDF, btnCancel})
        exportDialog.ShowDialog()
    End Sub

    ' Export to CSV
    Private Sub ExportToCSV()
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "CSV files (*.csv)|*.csv"
            saveDialog.FileName = "SensorData_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Dim csv As New StringBuilder()

                ' Add headers
                Dim headers As New List(Of String)()
                For Each column As DataGridViewColumn In DataGridView1.Columns
                    If column.Visible Then
                        headers.Add("""" & column.HeaderText & """")
                    End If
                Next
                csv.AppendLine(String.Join(",", headers))

                ' Add data rows
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        Dim values As New List(Of String)()
                        For Each cell As DataGridViewCell In row.Cells
                            If DataGridView1.Columns(cell.ColumnIndex).Visible Then
                                Dim value As String = If(cell.Value?.ToString(), "")
                                values.Add("""" & value.Replace("""", """""") & """")
                            End If
                        Next
                        csv.AppendLine(String.Join(",", values))
                    End If
                Next

                File.WriteAllText(saveDialog.FileName, csv.ToString())
                MessageBox.Show("Data exported successfully to CSV!")
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting to CSV: " & ex.Message)
        End Try
    End Sub

    ' Export to Excel
    Private Sub ExportToExcel()
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
            saveDialog.FileName = "SensorData_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Dim xlApp As New Excel.Application()
                Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Add()
                Dim xlWorksheet As Excel.Worksheet = CType(xlWorkbook.Sheets(1), Excel.Worksheet)

                ' Add headers
                Dim colIndex As Integer = 1
                For Each column As DataGridViewColumn In DataGridView1.Columns
                    If column.Visible Then
                        xlWorksheet.Cells(1, colIndex) = column.HeaderText
                        xlWorksheet.Cells(1, colIndex).Font.Bold = True
                        colIndex += 1
                    End If
                Next

                ' Add data
                Dim rowIndex As Integer = 2
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        colIndex = 1
                        For Each cell As DataGridViewCell In row.Cells
                            If DataGridView1.Columns(cell.ColumnIndex).Visible Then
                                xlWorksheet.Cells(rowIndex, colIndex) = cell.Value?.ToString()
                                colIndex += 1
                            End If
                        Next
                        rowIndex += 1
                    End If
                Next

                xlWorksheet.Columns.AutoFit()
                xlWorkbook.SaveAs(saveDialog.FileName)
                xlWorkbook.Close()
                xlApp.Quit()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Data exported successfully to Excel!")
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting to Excel: " & ex.Message & vbCrLf & "Make sure Microsoft Excel is installed.")
        End Try
    End Sub

    ' Export to Word
    Private Sub ExportToWord()
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Word documents (*.docx)|*.docx"
            saveDialog.FileName = "SensorData_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Dim wordApp As New Word.Application()
                Dim doc As Word.Document = wordApp.Documents.Add()

                ' Add title
                Dim title As Word.Paragraph = doc.Content.Paragraphs.Add()
                title.Range.Text = "Sensor Data Report" & vbCrLf & "Generated: " & DateTime.Now.ToString() & vbCrLf & vbCrLf
                title.Range.Font.Bold = True
                title.Range.Font.Size = 14

                ' Create table
                Dim visibleColumns As Integer = DataGridView1.Columns.Cast(Of DataGridViewColumn)().Count(Function(c) c.Visible)
                Dim dataRows As Integer = DataGridView1.Rows.Cast(Of DataGridViewRow)().Count(Function(r) Not r.IsNewRow)

                Dim table As Word.Table = doc.Tables.Add(doc.Content, dataRows + 1, visibleColumns)
                table.Borders.Enable = True

                ' Add headers
                Dim colIndex As Integer = 1
                For Each column As DataGridViewColumn In DataGridView1.Columns
                    If column.Visible Then
                        table.Cell(1, colIndex).Range.Text = column.HeaderText
                        table.Cell(1, colIndex).Range.Font.Bold = True
                        colIndex += 1
                    End If
                Next

                ' Add data
                Dim rowIndex As Integer = 2
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        colIndex = 1
                        For Each cell As DataGridViewCell In row.Cells
                            If DataGridView1.Columns(cell.ColumnIndex).Visible Then
                                table.Cell(rowIndex, colIndex).Range.Text = cell.Value?.ToString()
                                colIndex += 1
                            End If
                        Next
                        rowIndex += 1
                    End If
                Next

                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
                doc.SaveAs2(saveDialog.FileName)
                doc.Close()
                wordApp.Quit()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(table)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)

                MessageBox.Show("Data exported successfully to Word!")
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting to Word: " & ex.Message & vbCrLf & "Make sure Microsoft Word is installed.")
        End Try
    End Sub

    ' Export to PDF
    Private Sub ExportToPDF()
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "PDF files (*.pdf)|*.pdf"
            saveDialog.FileName = "SensorData_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pdf"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Dim doc As New iTextSharp.text.Document(PageSize.A4, 10, 10, 10, 10)
                Dim writer As PdfWriter = PdfWriter.GetInstance(doc, New FileStream(saveDialog.FileName, FileMode.Create))
                doc.Open()

                ' Add title
                Dim titleFont As New iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD)
                Dim title As New Paragraph("Sensor Data Report", titleFont)
                title.Alignment = Element.ALIGN_CENTER
                doc.Add(title)
                doc.Add(New Paragraph(" "))
                doc.Add(New Paragraph("Generated: " & DateTime.Now.ToString()))
                doc.Add(New Paragraph(" "))

                ' Create table
                Dim visibleColumns As Integer = DataGridView1.Columns.Cast(Of DataGridViewColumn)().Count(Function(c) c.Visible)
                Dim table As New PdfPTable(visibleColumns)
                table.WidthPercentage = 100

                ' Add headers
                For Each column As DataGridViewColumn In DataGridView1.Columns
                    If column.Visible Then
                        Dim headerCell As New PdfPCell(New Phrase(column.HeaderText))
                        headerCell.BackgroundColor = BaseColor.LIGHT_GRAY
                        headerCell.HorizontalAlignment = Element.ALIGN_CENTER
                        table.AddCell(headerCell)
                    End If
                Next

                ' Add data
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        For Each cell As DataGridViewCell In row.Cells
                            If DataGridView1.Columns(cell.ColumnIndex).Visible Then
                                table.AddCell(cell.Value?.ToString())
                            End If
                        Next
                    End If
                Next

                doc.Add(table)
                doc.Close()
                writer.Close()

                MessageBox.Show("Data exported successfully to PDF!")
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting to PDF: " & ex.Message & vbCrLf & "Make sure iTextSharp library is installed.")
        End Try
    End Sub

    Private Sub form2Btn_Click(sender As Object, e As EventArgs) Handles form2Btn.Click
        Dim f2 As New Form2()
        f2.ShowDialog(Me)
    End Sub
End Class