Imports System.IO
Imports System.Text.RegularExpressions
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class Form1
    Private filePath As String

    ' 窗体加载事件
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 设置 EPPlus 许可上下文
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' 初始化组件
        InitializeComponent()

        ' 设置窗体的初始设置
        Me.ClientSize = New Size(800, 600)  ' 根据需要调整窗体大小
        Me.Text = "Excel 数据处理与比对"
    End Sub

    ' 按钮点击事件
    Public Sub MainMethod(file1 As String, file2 As String)

        '' 异步处理以避免界面冻结
        Try
            Using package1 As New ExcelPackage(New FileInfo(file1)), package2 As New ExcelPackage(New FileInfo(file2))
                Dim worksheetN601 As ExcelWorksheet = package1.Workbook.Worksheets("N601")
                ' 获取第一个文件工作表
                Dim wsTarget As ExcelWorksheet = package2.Workbook.Worksheets("科目汇总表查询.xlsx") ' 获取第二个文件工作表

                ' 调用 ProcessExcelFile 方法处理第一个文件
                ProcessExcelFile(worksheetN601, wsTarget, "C") ' 假设要处理第一个工作表并指定目标列为"A"
                SaveProcessedFile(package1, file1)
                ' 调用 ProcessExcelFile 方法处理第二个文件
            End Using



        Catch ex As Exception
            MessageBox.Show("发生错误: " & ex.Message)
        Finally
            ' 重新启用按钮
            btnCompare.Enabled = True
        End Try
    End Sub

    ' 处理Excel文件的公共方法
    Public Sub ProcessExcelFile(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, targetColumn As String)
        ' 检查工作表是否存在
        If wsSource Is Nothing OrElse wsTarget Is Nothing Then
            Throw New Exception("源工作表或目标工作表不存在！")
        End If
        If wsSource.Dimension Is Nothing Then
            Throw New Exception("wsSource 工作表没有有效范围")
        End If

        If wsTarget.Dimension Is Nothing Then
            Throw New Exception("wsTarget 工作表没有有效范围")
        End If

        ' 获取使用区域
        Dim sourceEndRow As Integer = Math.Max(11, If(wsSource.Dimension?.End.Row, 11)) ' 确保 >= 11
        Dim targetEndRow As Integer = Math.Max(4, If(wsTarget.Dimension?.End.Row, 4))   ' 确保 >= 4

        Dim rangeSource As ExcelRange = wsSource.Cells($"A11:A{sourceEndRow}")
        Dim rangeTarget As ExcelRange = wsTarget.Cells($"A4:A{targetEndRow}")

        ' 清理目标工作表 A 列的值（去除空格）
        CleanColumn(wsTarget, "A")

        ' 遍历源工作表 A 列
        Dim checkValue As Integer = 221101 ' 根据需要设置checkValue的值
        For Each currentCell In rangeSource
            ' 获取目标列名称（假设为H或M，根据实际情况调整）
            ' 如果目标列为空，则跳过当前单元格
            Dim columnIndex As Integer = ColumnLetterToNumber(targetColumn) ' 返回列号
            If String.IsNullOrEmpty(targetColumn) OrElse String.IsNullOrWhiteSpace(wsSource.Cells(currentCell.Start.Row, columnIndex).Text) Then
                Continue For
            End If

            ' 清理当前源单元格 A 列的值，去除空格并去掉括号
            Dim cleanedValue As String = RemoveBrackets(Trim(currentCell.Text))

            ' 根据源工作表 A 列的值，确定需要进行匹配的目标值（可能有多个匹配项）
            Dim possibleMatches As String() = GetPossibleMatches(cleanedValue)

            ' 初始化累加计算结果
            Dim totalCalculatedValue As Double = 0
            Dim matchFound As Boolean = False ' 用来标记是否找到匹配项

            ' 根据工作表名称和目标列应用不同的匹配逻辑
            If IsSpecialSheet(wsSource.Name, targetColumn) Then
                totalCalculatedValue = ProcessSpecialSheets(wsSource, wsTarget, currentCell.Start.Row, targetColumn, possibleMatches)
                matchFound = totalCalculatedValue > 0
            Else
                totalCalculatedValue = ProcessGeneralSheets(wsSource, wsTarget, currentCell.Start.Row, targetColumn, possibleMatches, checkValue)
                matchFound = totalCalculatedValue > 0
            End If

            ' 如果工作表名称是N601，处理V和AD列数据
            If wsSource.Name = "N601" Then
                totalCalculatedValue = MergeAndCenterCells(wsSource, currentCell.Start.Row, targetColumn, totalCalculatedValue)
            End If

            ' 根据累加的结果设置颜色
            Dim valueSource As Double = 0
            If Double.TryParse(wsSource.Cells(currentCell.Start.Row, columnIndex).Text, valueSource) Then
                If matchFound Then
                    If valueSource = totalCalculatedValue Then
                        SetCellColor(wsSource, currentCell.Start.Row, columnIndex, Color.FromArgb(0, 190, 140)) ' 一致标绿
                    Else
                        SetCellColor(wsSource, currentCell.Start.Row, columnIndex, Color.Red) ' 不一致标红
                    End If
                Else
                    ' 未匹配到，标黄
                    SetCellColor(wsSource, currentCell.Start.Row, columnIndex, Color.FromArgb(250, 240, 230)) ' 标黄
                End If
            End If
        Next
    End Sub

    ' 获取指定名称的工作表
    Public Function GetWorksheet(package As ExcelPackage, sheetName As String) As ExcelWorksheet
        Return package.Workbook.Worksheets(sheetName)
    End Function

    ' 获取目标列名称，根据工作表名称或其他条件
    Public Function GetTargetColumn(sheetName As String) As String
        Select Case sheetName
            Case "N607-3", "N607-1", "N607-4"
                Return "H"
            Case "N607-5"
                Return "M"
            Case Else
                Return "H" ' 默认目标列，可根据需要调整
        End Select
    End Function

    ' 判断是否为特殊工作表
    Public Function IsSpecialSheet(sheetName As String, targetColumn As String) As Boolean
        If (sheetName = "N607-3" OrElse sheetName = "N607-1" OrElse sheetName = "N607-4") AndAlso targetColumn = "H" Then
            Return True
        End If
        If (sheetName = "N607-3" OrElse sheetName = "N607-4") AndAlso targetColumn = "M" Then
            Return True
        End If
        Return False
    End Function

    ' 获取可能的匹配项
    Public Function GetPossibleMatches(cleanedValue As String) As String()
        Select Case cleanedValue
            Case "国网内蒙古东部电力有限公司物资事业部"
                Return {"国网内蒙古东部电力有限公司物资分公司", "国网内蒙古东部电力招标有限公司"}
            Case "国网内蒙古东部电力有限公司数字化事业部"
                Return {"国网内蒙古东部电力有限公司信息通信分公司"}
            Case "国网内蒙古东部电力有限公司经济技术研究院"
                Return {"国网内蒙古东部电力有限公司经济技术研究院", "国网内蒙古东部电力设计有限公司"}
            Case Else
                Return {cleanedValue}
        End Select
    End Function

    ' 清理指定列的空格
    Public Sub CleanColumn(ws As ExcelWorksheet, column As String)
        'Dim range As ExcelRange = ws.Cells($"{column}4", $"{column}{ws.Dimension.End.Row}")
        Dim targetEndRow As Integer = Math.Max(4, If(ws.Dimension?.End.Row, 4))   ' 确保 >= 4
        Dim range As ExcelRange = ws.Cells($"{column}4:{column}{targetEndRow}") ' 修正为指定列
        Dim cleanedData As Object(,) = range.Value

        ' 检查范围内是否有数据
        If range.Value Is Nothing Then
            Throw New Exception($"的范围 {column}4 到 {column}{targetEndRow} 内没有数据")
        End If

        ' 遍历指定列的数据
        For i As Integer = 0 To cleanedData.GetLength(0) - 1
            If cleanedData(i, 0) IsNot Nothing AndAlso TypeOf cleanedData(i, 0) Is String Then
                cleanedData(i, 0) = Trim(cleanedData(i, 0).ToString())
            End If
        Next

        ' 输出调试信息
        Console.WriteLine($"的数据已成功清理，处理范围: {column}4 到 {column}{targetEndRow}")

        range.Value = cleanedData
    End Sub

    ' 处理特殊工作表的匹配和计算
    Public Function ProcessSpecialSheets(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, sourceRow As Integer, targetColumn As String, possibleMatches As String()) As Double
        Dim totalCalculatedValue As Double = 0
        Dim columnA As Integer = CInt(ColumnLetterToNumber("A"))
        Dim columnB As Integer = CInt(ColumnLetterToNumber("B"))
        Dim columnE As Integer = CInt(ColumnLetterToNumber("E"))
        Dim columnF As Integer = CInt(ColumnLetterToNumber("F"))
        Dim columnI As Integer = CInt(ColumnLetterToNumber("I"))

        ' 遍历可能的匹配项
        For Each match In possibleMatches
            ' 遍历目标工作表的 A 列，查找所有匹配的项
            For i As Integer = 1 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = If(wsTarget.Cells(i, ColumnLetterToNumber("A")).Value IsNot Nothing, Trim(wsTarget.Cells(i, ColumnLetterToNumber("A")).Value.ToString()), "")
                Dim targetCValue As String = If(wsTarget.Cells(i, ColumnLetterToNumber("C")).Value IsNot Nothing, wsTarget.Cells(i, ColumnLetterToNumber("C")).Value.ToString(), "")

                ' 如果 C 列包含 "劳动保护费" 并且 A 列匹配
                If targetCValue.Contains("劳动保护费") AndAlso targetCellValue = match Then
                    Dim eVal As Double = If(IsNumeric(wsTarget.Cells(i, "E").Value), CDbl(wsTarget.Cells(i, "E").Value), 0)
                    Dim fVal As Double = If(IsNumeric(wsTarget.Cells(i, "F").Value), CDbl(wsTarget.Cells(i, "F").Value), 0)
                    Dim iVal As Double = If(IsNumeric(wsTarget.Cells(i, "I").Value), CDbl(wsTarget.Cells(i, "I").Value), 0)

                    totalCalculatedValue += (eVal + fVal - iVal)
                End If
            Next
        Next

        Return totalCalculatedValue
    End Function

    ' 将列名转换为列号
    Function ColumnLetterToNumber(column As String) As Integer
        ' 验证输入是否合法
        If String.IsNullOrEmpty(column) OrElse column.Length > 3 OrElse Not Regex.IsMatch(column, "^[A-Z]+$") Then
            Throw New ArgumentException($"无效的列名: {column}")
        End If

        Dim colNumber As Integer = 0
        For Each ch As Char In column.ToUpper()
            colNumber = colNumber * 26 + (Asc(ch) - Asc("A") + 1)
        Next

        ' 确保列号在 Excel 支持范围内
        If colNumber > 16384 Then
            Throw New ArgumentOutOfRangeException($"列号 {colNumber} 超出了 Excel 的最大支持范围 (XFD)")
        End If

        Return colNumber
    End Function

    ' 处理一般工作表的匹配和计算
    Public Function ProcessGeneralSheets(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, sourceRow As Integer, targetColumn As String, possibleMatches As String(), checkValue As Integer) As Decimal
        Dim totalCalculatedValue As Decimal = 0D

        ' 将列名转换为列号
        Dim columnA As Integer = ColumnLetterToNumber("A")
        Dim columnB As Integer = ColumnLetterToNumber("B")
        Dim columnE As Integer = ColumnLetterToNumber("E")
        Dim columnF As Integer = ColumnLetterToNumber("F")
        Dim columnI As Integer = ColumnLetterToNumber("I")

        ' 获取源工作表中 sourceRow 的 A 列值
        Dim sourceValue As String = If(wsSource.Cells(sourceRow, columnA).Value IsNot Nothing, Trim(wsSource.Cells(sourceRow, columnA).Value.ToString()), "")

        ' 优化匹配项为 HashSet
        Dim matchesSet As New HashSet(Of String)(possibleMatches)

        ' 遍历目标工作表的行
        For i As Integer = 4 To wsTarget.Dimension.End.Row
            ' 获取目标工作表中当前行的 A 列值
            Dim targetCellValue As String = If(wsTarget.Cells(i, columnA).Value IsNot Nothing, Trim(wsTarget.Cells(i, columnA).Value.ToString()), "")

            ' 如果匹配项和检查值符合条件
            If matchesSet.Contains(targetCellValue) Then
                ' 尝试将 B 列的值转换为 Long 类型
                Dim targetBValue As Long
                If Not Long.TryParse(wsTarget.Cells(i, columnB).Value?.ToString(), targetBValue) Then
                    targetBValue = 0L
                    ' 记录无效的 B 列数值
                End If

                If targetBValue = checkValue Then
                    Dim eVal As Decimal = 0D
                    Dim fVal As Decimal = 0D
                    Dim iVal As Decimal = 0D

                    ' 安全转换 E 列数值
                    If Not Decimal.TryParse(wsTarget.Cells(i, columnE).Value?.ToString(), eVal) Then
                        eVal = 0D
                    End If

                    ' 安全转换 F 列数值
                    If Not Decimal.TryParse(wsTarget.Cells(i, columnF).Value?.ToString(), fVal) Then
                        fVal = 0D
                    End If

                    ' 安全转换 I 列数值
                    If Not Decimal.TryParse(wsTarget.Cells(i, columnI).Value?.ToString(), iVal) Then
                        iVal = 0D
                    End If

                    ' 检查异常数据，包括负值
                    If Math.Abs(eVal) > 1000000000000000D OrElse Math.Abs(fVal) > 1000000000000000D OrElse Math.Abs(iVal) > 1000000000000000D Then
                        Throw New Exception($"异常数据: E={eVal}, F={fVal}, I={iVal} (行 {i})")
                    End If

                    ' 计算目标行的结果：E + F - I
                    Dim calculatedValue As Decimal
                    Try
                        calculatedValue = eVal + fVal - iVal
                    Catch ex As OverflowException
                        Throw New Exception($"计算溢出: E={eVal}, F={fVal}, I={iVal} (行 {i})")
                    End Try

                    ' 如果参数 wsSource.Name = "N601" 调用处理AD列和V列数据
                    If wsSource.Name = "N601" Then
                        calculatedValue = MergeAndCenterCells(wsSource, sourceRow, targetColumn, calculatedValue)
                    End If

                    ' 累加计算结果
                    totalCalculatedValue += calculatedValue
                End If
            End If
        Next

        Return totalCalculatedValue
    End Function

    ' 示例的日志记录方法
    Public Sub WriteLog(message As String)
        Dim logFilePath As String = Path.Combine(Path.GetDirectoryName(filePath), "ProcessLog.txt")
        SyncLock Me
            Using writer As New StreamWriter(logFilePath, True)
                writer.WriteLine($"{DateTime.Now}: {message}")
            End Using
        End SyncLock
    End Sub

    ' 合并并居中指定列的数据
    Public Function MergeAndCenterCells(wsSource As ExcelWorksheet, sourceRow As Integer, targetColumn As String, valueSource As Double) As Double
        Dim valueColumn1 As Double = 0
        Dim valueColumn2 As Double = 0
        ' 将列名转换为列号
        Dim columnV As Integer = ColumnLetterToNumber("V")
        Dim columnW As Integer = ColumnLetterToNumber("W")
        Dim columnAD As Integer = ColumnLetterToNumber("AD")
        Dim columnAE As Integer = ColumnLetterToNumber("AE")

        If targetColumn = "V" Then
            valueColumn1 = If(IsNumeric(wsSource.Cells(sourceRow, columnV).Value), CDbl(wsSource.Cells(sourceRow, columnV).Value), 0)
            valueColumn2 = If(IsNumeric(wsSource.Cells(sourceRow, columnW).Value), CDbl(wsSource.Cells(sourceRow, columnW).Value), 0)
            valueSource = valueColumn1 + valueColumn2

            ' 合并 V 列和 W 列的单元格，并居中显示
            Using range As ExcelRange = wsSource.Cells(sourceRow, columnV, sourceRow, columnW)
                range.Merge = True
                range.Value = valueSource
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center
            End Using
        End If

        If targetColumn = "AD" Then
            valueColumn1 = If(IsNumeric(wsSource.Cells(sourceRow, columnAD).Value), CDbl(wsSource.Cells(sourceRow, columnAD).Value), 0)
            valueColumn2 = If(IsNumeric(wsSource.Cells(sourceRow, columnAE).Value), CDbl(wsSource.Cells(sourceRow, columnAE).Value), 0)
            valueSource = valueColumn1 + valueColumn2

            ' 合并 AD 列和 AE 列的单元格，并居中显示
            Using range As ExcelRange = wsSource.Cells(sourceRow, columnAD, sourceRow, columnAE)
                range.Merge = True
                range.Value = valueSource
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center
            End Using
        End If

        Return valueSource
    End Function

    ' 设置单元格颜色的公共方法
    Public Sub SetCellColor(ws As ExcelWorksheet, row As Integer, column As Integer, color As Color)
        ws.Cells(row, column).Style.Fill.PatternType = ExcelFillStyle.Solid
        ws.Cells(row, column).Style.Fill.BackgroundColor.SetColor(color)
    End Sub

    ' 移除字符串中的括号和空格
    Public Function RemoveBrackets(originalString As String) As String
        Dim cleanedString As String = originalString.Replace(" ", "") _
                                                 .Replace("【", "") _
                                                 .Replace("】", "") _
                                                 .Replace("(", "") _
                                                 .Replace(")", "") _
                                                 .Replace("[", "") _
                                                 .Replace("]", "")
        Return cleanedString
    End Function

    ' 提供保存文件对话框并保存处理后的文件
    Public Sub SaveProcessedFile(package As ExcelPackage, originalFilePath As String)
        ' 初始化 SaveFileDialog
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Excel 文件|*.xlsx"
        saveFileDialog1.Title = "保存处理结果"
        saveFileDialog1.FileName = Path.GetFileNameWithoutExtension(originalFilePath) & "_处理结果.xlsx"

        ' 弹出对话框并等待用户操作
        Dim dialogResult As DialogResult = saveFileDialog1.ShowDialog()

        If dialogResult = DialogResult.OK Then
            ' 保存文件到用户选择的路径
            Dim savePath As String = saveFileDialog1.FileName
            package.SaveAs(New FileInfo(savePath)) ' 保存文件
            MessageBox.Show("文件已保存: " & savePath, "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' 如果用户取消了保存操作
            MessageBox.Show("文件保存被取消。", "取消", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class
