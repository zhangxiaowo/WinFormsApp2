Partial Class Form1
    Inherits Form

    Private components As System.ComponentModel.IContainer
    Private WithEvents btnCompare As Button
    Private openFileDialog1 As OpenFileDialog
    Private openFileDialog2 As OpenFileDialog
    Private txtFile1 As TextBox
    Private txtFile2 As TextBox
    Private btnBrowse1 As Button
    Private btnBrowse2 As Button
    Private Label1 As Label
    Private Label2 As Label

    ' 窗体初始化代码
    Private Sub InitializeComponent()
        btnCompare = New Button()
        openFileDialog1 = New OpenFileDialog()
        openFileDialog2 = New OpenFileDialog()
        txtFile1 = New TextBox()
        txtFile2 = New TextBox()
        btnBrowse1 = New Button()
        btnBrowse2 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        SuspendLayout()

        ' 
        ' btnCompare
        ' 
        btnCompare.BackColor = Color.FromArgb(64, 158, 255) ' 现代化蓝色
        btnCompare.ForeColor = Color.White
        btnCompare.FlatStyle = FlatStyle.Flat
        btnCompare.FlatAppearance.BorderSize = 0
        btnCompare.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnCompare.Location = New Point(292, 360)
        btnCompare.Name = "btnCompare"
        btnCompare.Size = New Size(200, 60)
        btnCompare.TabIndex = 0
        btnCompare.Text = "开始校验"
        btnCompare.UseVisualStyleBackColor = True
        AddHandler btnCompare.Click, AddressOf btnCompare_Click
        ' 鼠标悬停效果
        AddHandler btnCompare.MouseEnter, Sub(sender, e) btnCompare.BackColor = Color.FromArgb(0, 102, 204)
        AddHandler btnCompare.MouseLeave, Sub(sender, e) btnCompare.BackColor = Color.FromArgb(64, 158, 255)

        ' 
        ' txtFile1
        ' 
        txtFile1.BackColor = Color.White
        txtFile1.BorderStyle = BorderStyle.None
        txtFile1.Font = New Font("Segoe UI", 12)
        txtFile1.Location = New Point(154, 100)
        txtFile1.Name = "txtFile1"
        txtFile1.ReadOnly = True
        txtFile1.Size = New Size(500, 34)
        txtFile1.TabIndex = 1
        txtFile1.ForeColor = Color.FromArgb(80, 80, 80)
        txtFile1.BackColor = Color.FromArgb(248, 249, 250)
        txtFile1.Padding = New Padding(10)
        txtFile1.TextAlign = HorizontalAlignment.Left

        ' 
        ' txtFile2
        ' 
        txtFile2.BackColor = Color.White
        txtFile2.BorderStyle = BorderStyle.None
        txtFile2.Font = New Font("Segoe UI", 12)
        txtFile2.Location = New Point(154, 153)
        txtFile2.Name = "txtFile2"
        txtFile2.ReadOnly = True
        txtFile2.Size = New Size(500, 34)
        txtFile2.TabIndex = 2
        txtFile2.ForeColor = Color.FromArgb(80, 80, 80)
        txtFile2.BackColor = Color.FromArgb(248, 249, 250)
        txtFile2.Padding = New Padding(10)
        txtFile2.TextAlign = HorizontalAlignment.Left

        ' 
        ' btnBrowse1
        ' 
        btnBrowse1.BackColor = Color.FromArgb(64, 158, 255) ' 现代化蓝色
        btnBrowse1.FlatStyle = FlatStyle.Flat
        btnBrowse1.ForeColor = Color.White
        btnBrowse1.Font = New Font("Segoe UI", 12)
        btnBrowse1.Location = New Point(675, 100)
        btnBrowse1.Name = "btnBrowse1"
        btnBrowse1.Size = New Size(91, 34)
        btnBrowse1.TabIndex = 3
        btnBrowse1.Text = "浏览"
        btnBrowse1.UseVisualStyleBackColor = True
        AddHandler btnBrowse1.Click, AddressOf btnBrowse1_Click
        ' 鼠标悬停效果
        AddHandler btnBrowse1.MouseEnter, Sub(sender, e) btnBrowse1.BackColor = Color.FromArgb(0, 102, 204)
        AddHandler btnBrowse1.MouseLeave, Sub(sender, e) btnBrowse1.BackColor = Color.FromArgb(64, 158, 255)

        ' 
        ' btnBrowse2
        ' 
        btnBrowse2.BackColor = Color.FromArgb(64, 158, 255) ' 现代化蓝色
        btnBrowse2.FlatStyle = FlatStyle.Flat
        btnBrowse2.ForeColor = Color.White
        btnBrowse2.Font = New Font("Segoe UI", 12)
        btnBrowse2.Location = New Point(675, 153)
        btnBrowse2.Name = "btnBrowse2"
        btnBrowse2.Size = New Size(91, 34)
        btnBrowse2.TabIndex = 4
        btnBrowse2.Text = "浏览"
        btnBrowse2.UseVisualStyleBackColor = True
        AddHandler btnBrowse2.Click, AddressOf btnBrowse2_Click
        ' 鼠标悬停效果
        AddHandler btnBrowse2.MouseEnter, Sub(sender, e) btnBrowse2.BackColor = Color.FromArgb(0, 102, 204)
        AddHandler btnBrowse2.MouseLeave, Sub(sender, e) btnBrowse2.BackColor = Color.FromArgb(64, 158, 255)

        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 14)
        Label1.ForeColor = Color.FromArgb(50, 50, 50)
        Label1.Location = New Point(36, 100)
        Label1.Name = "Label1"
        Label1.Size = New Size(96, 28)
        Label1.TabIndex = 5
        Label1.Text = "人资表："

        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 14)
        Label2.ForeColor = Color.FromArgb(50, 50, 50)
        Label2.Location = New Point(36, 153)
        Label2.Name = "Label2"
        Label2.Size = New Size(96, 28)
        Label2.TabIndex = 6
        Label2.Text = "科目表："

        ' 
        ' Form1
        ' 
        ClientSize = New Size(800, 700)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(btnCompare)
        Controls.Add(txtFile1)
        Controls.Add(txtFile2)
        Controls.Add(btnBrowse1)
        Controls.Add(btnBrowse2)
        Name = "Form1"
        Text = "Excel 比对"
        BackColor = Color.White
        Font = New Font("Segoe UI", 12)
        ResumeLayout(False)
        PerformLayout()
    End Sub

    ' 按钮点击事件：选择第一个文件
    Private Sub btnBrowse1_Click(sender As Object, e As EventArgs)
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            txtFile1.Text = openFileDialog1.FileName
        End If
    End Sub

    ' 按钮点击事件：选择第二个文件
    Private Sub btnBrowse2_Click(sender As Object, e As EventArgs)
        If openFileDialog2.ShowDialog() = DialogResult.OK Then
            txtFile2.Text = openFileDialog2.FileName
        End If
    End Sub

    ' 按钮点击事件：执行比较操作
    Private Sub btnCompare_Click(sender As Object, e As EventArgs)
        Dim file1 As String = txtFile1.Text
        Dim file2 As String = txtFile2.Text

        ' 检查用户是否选择了文件
        If String.IsNullOrEmpty(file1) Or String.IsNullOrEmpty(file2) Then
            MessageBox.Show("请确保选择两个文件进行比对！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        btnCompare.Enabled = False
        MainMethod(file1, file2)
    End Sub
End Class
