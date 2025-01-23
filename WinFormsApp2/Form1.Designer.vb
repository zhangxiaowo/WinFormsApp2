Partial Class Form1
    Inherits Form

    Private components As System.ComponentModel.IContainer
    Private WithEvents btnCompare As Button
    Private openFileDialog1 As OpenFileDialog

    ' 窗体初始化代码
    Private Sub InitializeComponent()
        Me.btnCompare = New Button()
        Me.openFileDialog1 = New OpenFileDialog()
        Me.SuspendLayout()

        ' btnCompare 按钮的设置
        Me.btnCompare.Location = New Point(280, 120)  ' 设置按钮位置
        Me.btnCompare.Name = "btnCompare"
        Me.btnCompare.Size = New Size(200, 60)  ' 设置按钮大小
        Me.btnCompare.TabIndex = 0
        Me.btnCompare.Text = "开始比对"
        Me.btnCompare.UseVisualStyleBackColor = True

        ' openFileDialog1
        Me.openFileDialog1.Filter = "Excel 文件 (*.xlsx)|*.xlsx"

        ' 将按钮添加到窗体
        Me.Controls.Add(Me.btnCompare)

        ' 窗体的基本设置
        Me.ClientSize = New Size(500, 500)  ' 设置窗体大小
        Me.Name = "Form1"
        Me.Text = "Excel 比对"
        Me.ResumeLayout(False)
    End Sub
End Class
