<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.form2Btn = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.refreshBtn = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Window
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(13, 77)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(637, 662)
        Me.DataGridView1.TabIndex = 0
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.SystemColors.Highlight
        Me.btnRefresh.Location = New System.Drawing.Point(13, 761)
        Me.btnRefresh.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(166, 59)
        Me.btnRefresh.TabIndex = 1
        Me.btnRefresh.Text = "Refresh BME"
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(13, 9)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(637, 59)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "BME Sensor Data"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnExport
        '
        Me.btnExport.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnExport.Location = New System.Drawing.Point(464, 761)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(186, 59)
        Me.btnExport.TabIndex = 3
        Me.btnExport.Text = "Export File"
        Me.btnExport.UseVisualStyleBackColor = False
        '
        'form2Btn
        '
        Me.form2Btn.BackColor = System.Drawing.SystemColors.Info
        Me.form2Btn.Location = New System.Drawing.Point(1273, 761)
        Me.form2Btn.Name = "form2Btn"
        Me.form2Btn.Size = New System.Drawing.Size(237, 59)
        Me.form2Btn.TabIndex = 5
        Me.form2Btn.Text = "Load Camera Data"
        Me.form2Btn.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Info
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Location = New System.Drawing.Point(657, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(853, 59)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "MPU Sensor Data"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGridView2
        '
        Me.DataGridView2.BackgroundColor = System.Drawing.SystemColors.Window
        Me.DataGridView2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(657, 77)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersWidth = 51
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(856, 663)
        Me.DataGridView2.TabIndex = 10
        '
        'refreshBtn
        '
        Me.refreshBtn.BackColor = System.Drawing.SystemColors.Highlight
        Me.refreshBtn.Location = New System.Drawing.Point(657, 761)
        Me.refreshBtn.Name = "refreshBtn"
        Me.refreshBtn.Size = New System.Drawing.Size(188, 59)
        Me.refreshBtn.TabIndex = 12
        Me.refreshBtn.Text = "Refresh MPU"
        Me.refreshBtn.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1522, 841)
        Me.Controls.Add(Me.refreshBtn)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.form2Btn)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.DataGridView1)
        Me.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "JonBalSat"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnRefresh As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents btnExport As Button
    Friend WithEvents form2Btn As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents refreshBtn As Button
End Class
