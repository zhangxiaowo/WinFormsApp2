Imports System.Data.Common
Imports System.IO
Imports System.Text.RegularExpressions
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class Form1
    Private filePath As String

    ' 窗体加载事件1
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
                Dim worksheetN602 As ExcelWorksheet = package1.Workbook.Worksheets("N602")
                Dim worksheetN606 As ExcelWorksheet = package1.Workbook.Worksheets("N606")
                Dim worksheetN607_1 As ExcelWorksheet = package1.Workbook.Worksheets("N607-1")
                Dim worksheetN607_3 As ExcelWorksheet = package1.Workbook.Worksheets("N607-3")
                Dim worksheetN607_4 As ExcelWorksheet = package1.Workbook.Worksheets("N607-4")
                Dim checkValue221101 As Long = 221101   '工资
                Dim checkValue221102 As Long = 221102   '职工福利
                Dim checkValue221103 As Long = 221103   '工会经费
                Dim checkValue221104 As Long = 221104   '职工教育经费
                Dim checkValue221105 As Long = 221105   '辞退福利
                Dim checkValue221109 As Long = 221109   '住房公积金
                Dim checkValue221113 As Long = 221113   '劳务派遣费
                Dim checkValue221114 As Long = 221114   '临时用工薪酬
                Dim checkValue221115 As Long = 221115   '农电工用工薪酬
                Dim checkValue221117 As Long = 221117   '特殊工种保险费
                Dim checkValue221118 As Long = 221118   '社会保险费
                ' 获取第一个文件工作表
                Dim wsTarget As ExcelWorksheet = package2.Workbook.Worksheets("科目汇总表查询.xlsx") ' 获取第二个文件工作表

                ' 调用 ProcessExcelFile 方法处理第一个文件
                'N601
                ProcessExcelFile(worksheetN601, wsTarget, "C", checkValue221101) ' 
                ProcessExcelFile(worksheetN601, wsTarget, "E", checkValue221101) ' 
                ProcessExcelFile(worksheetN601, wsTarget, "F", checkValue221101) '
                ProcessExcelFile(worksheetN601, wsTarget, "K", checkValue221101) '
                ProcessExcelFile(worksheetN601, wsTarget, "V", checkValue221101) '
                ProcessExcelFile(worksheetN601, wsTarget, "AD", checkValue221101) '
                'N607-1
                ProcessExcelFile(worksheetN607_1, wsTarget, "C", checkValue221101) ' 
                ProcessExcelFile(worksheetN607_1, wsTarget, "D", checkValue221118) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "E", checkValue221104) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "F", checkValue221103) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "G", checkValue221109) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "I", checkValue221102) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "K", checkValue221105) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "N", checkValue221117) '
                ProcessExcelFile(worksheetN607_1, wsTarget, "H", checkValue221101) '
                'N607-3
                ProcessExcelFile(worksheetN607_3, wsTarget, "C", checkValue221101) ' 
                ProcessExcelFile(worksheetN607_3, wsTarget, "D", checkValue221118) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "E", checkValue221104) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "F", checkValue221103) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "G", checkValue221109) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "I", checkValue221102) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "J", checkValue221105) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "M", checkValue221117) '
                ProcessExcelFile(worksheetN607_3, wsTarget, "H", checkValue221101) '
                'N607-4
                ProcessExcelFile(worksheetN607_4, wsTarget, "C", checkValue221101) ' 
                ProcessExcelFile(worksheetN607_4, wsTarget, "D", checkValue221118) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "E", checkValue221104) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "F", checkValue221103) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "G", checkValue221109) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "I", checkValue221102) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "J", checkValue221105) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "M", checkValue221117) '
                ProcessExcelFile(worksheetN607_4, wsTarget, "H", checkValue221101) '
                ProcessN602(worksheetN602, worksheetN601)
                ProcessN606CK(worksheetN606, worksheetN607_1， "C")
                ProcessN606CK(worksheetN606, worksheetN607_1, "K")
                ProcessN606EF(worksheetN607_3, worksheetN606, wsTarget, "E")
                ProcessN606EF(worksheetN607_4, worksheetN606, wsTarget, "F")
                ProcessN606EF(worksheetN607_3, worksheetN606, wsTarget, "V")
                ProcessN606EF(worksheetN607_4, worksheetN606, wsTarget, "AD")
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
    Public Sub ProcessExcelFile(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, targetColumn As String, checkValue As Long)
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
        For Each currentCell In rangeSource
            ' 获取目标列名称（假设为H或M，根据实际情况调整）
            ' 如果目标列为空，则跳过当前单元格
            Dim columnIndex As Integer = ColumnLetterToNumber(targetColumn) ' 返回列号
            Dim cell As ExcelRange = wsSource.Cells(currentCell.Start.Row, columnIndex)
            ' 清除当前单元格的注释
            If cell.Comment IsNot Nothing Then
                cell.Clear() ' 清除注释
            End If

            If ShouldSkipCell(currentCell, "A", wsSource) Then
                Continue For
            End If
            Dim valueSource As Double = 0
            ' 清理当前源单元格 A 列的值，去除空格并去掉括号
            Dim cleanedValue As String = RemoveBrackets(Trim(currentCell.Text))

            ' 根据源工作表 A 列的值，确定需要进行匹配的目标值（可能有多个匹配项）
            Dim possibleMatches As String() = GetPossibleMatches(cleanedValue)

            ' 初始化累加计算结果
            Dim totalCalculatedValue As Double = 0
            Dim matchFound As Boolean = False ' 用来标记是否找到匹配项

            ' 根据工作表名称和目标列应用不同的匹配逻辑
            Dim specialSheetResult As Integer = IsSpecialSheet(wsSource.Name, targetColumn)
            valueSource = GetCellValueOrDefault(wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber(targetColumn)))
            If specialSheetResult = 1 Then
                ' 返回 1 时进入当前方法

                totalCalculatedValue = ProcessSpecialSheets(wsSource, wsTarget, currentCell.Start.Row, targetColumn, possibleMatches)
                matchFound = True
            ElseIf specialSheetResult = 2 Then
                ' 返回 2 时进入另一个方法
                totalCalculatedValue = ProcessAnotherSpecialSheet(wsSource, wsTarget, currentCell.Start.Row, targetColumn, possibleMatches)
                matchFound = True
            Else
                ' 否则按照现有逻辑执行
                totalCalculatedValue = ProcessGeneralSheets(wsSource, wsTarget, currentCell.Start.Row, targetColumn, possibleMatches, checkValue)
                matchFound = True
            End If


            ' 根据累加的结果设置颜色

            If Double.TryParse(wsSource.Cells(currentCell.Start.Row, columnIndex).Text, valueSource) Then
                ' 如果工作表名称是N601，处理V和AD列数据
                If wsSource.Name = "N601" Then
                    valueSource = MergeAndCenterCells(wsSource, currentCell.Start.Row, targetColumn, valueSource)
                End If
                If matchFound Then
                    If Math.Abs(valueSource - totalCalculatedValue) < 3 Then
                        SetCellColor(wsSource, currentCell.Start.Row, columnIndex, Color.FromArgb(0, 190, 140)) ' 一致标绿
                    Else
                        SetCellColor(wsSource, currentCell.Start.Row, columnIndex, Color.Red) ' 不一致标红

                        wsSource.Cells(currentCell.Start.Row, columnIndex).AddComment("科目表中的数据为 " & totalCalculatedValue)
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
    Public Function IsSpecialSheet(sheetName As String, targetColumn As String) As Integer
        If (sheetName = "N607-3" OrElse sheetName = "N607-1" OrElse sheetName = "N607-4") AndAlso targetColumn = "H" Then
            Return 1
        End If
        If (sheetName = "N607-3" OrElse sheetName = "N607-4") AndAlso targetColumn = "M" Then
            Return 2
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
            Case "国网内蒙古东部电力有限公司本部"
                Return {"国网内蒙古东部电力有限公司机关财务处"}
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
        Dim columnC As Integer = CInt(ColumnLetterToNumber("C"))
        Dim columnE As Integer = CInt(ColumnLetterToNumber("E"))
        Dim columnF As Integer = CInt(ColumnLetterToNumber("F"))
        Dim columnI As Integer = CInt(ColumnLetterToNumber("I"))

        ' 遍历可能的匹配项
        For Each match In possibleMatches
            ' 遍历目标工作表的 A 列，查找所有匹配的项
            For i As Integer = 4 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = If(wsTarget.Cells(i, columnA).Value IsNot Nothing, Trim(wsTarget.Cells(i, columnA).Value.ToString()), "")
                Dim targetCValue As String = If(wsTarget.Cells(i, columnC).Value IsNot Nothing, wsTarget.Cells(i, columnC).Value.ToString(), "")

                ' 如果 C 列包含 "劳动保护费" 并且 A 列匹配
                If targetCValue.Contains("劳动保护费") AndAlso targetCellValue = match Then
                    Dim eVal As Double = If(IsNumeric(wsTarget.Cells(i, columnE).Value), CDbl(wsTarget.Cells(i, columnE).Value), 0)
                    Dim fVal As Double = If(IsNumeric(wsTarget.Cells(i, columnF).Value), CDbl(wsTarget.Cells(i, columnF).Value), 0)
                    Dim iVal As Double = If(IsNumeric(wsTarget.Cells(i, columnI).Value), CDbl(wsTarget.Cells(i, columnI).Value), 0)

                    totalCalculatedValue += (eVal + fVal - iVal)
                End If
            Next
        Next

        Return totalCalculatedValue
    End Function
    Public Function ProcessAnotherSpecialSheet(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, sourceRow As Integer, targetColumn As String, possibleMatches As String()) As Double
        Dim totalCalculatedValue As Double = 0
        Dim columnA As Integer = CInt(ColumnLetterToNumber("A"))
        Dim columnB As Integer = CInt(ColumnLetterToNumber("B"))
        Dim columnE As Integer = CInt(ColumnLetterToNumber("E"))
        Dim columnF As Integer = CInt(ColumnLetterToNumber("F"))
        Dim columnI As Integer = CInt(ColumnLetterToNumber("I"))

        ' 遍历可能的匹配项
        For Each match In possibleMatches
            ' 遍历目标工作表的 A 列，查找所有匹配的项
            For i As Integer = 4 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = If(wsTarget.Cells(i, columnA).Value IsNot Nothing, Trim(wsTarget.Cells(i, columnA).Value.ToString()), "")
                Dim targetBValue As String = If(wsTarget.Cells(i, columnB).Value IsNot Nothing, wsTarget.Cells(i, columnB).Value.ToString(), "")

                If wsTarget.Cells(i, columnB).Value = 221113 Or
                           wsTarget.Cells(i, columnB).Value = 221114 Or
                           wsTarget.Cells(i, columnB).Value = 221117 Then
                    Dim eVal As Double = If(IsNumeric(wsTarget.Cells(i, columnE).Value), CDbl(wsTarget.Cells(i, columnE).Value), 0)
                    Dim fVal As Double = If(IsNumeric(wsTarget.Cells(i, columnF).Value), CDbl(wsTarget.Cells(i, columnF).Value), 0)
                    Dim iVal As Double = If(IsNumeric(wsTarget.Cells(i, columnI).Value), CDbl(wsTarget.Cells(i, columnI).Value), 0)

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

            '' 合并 V 列和 W 列的单元格，并居中显示
            'Using range As ExcelRange = wsSource.Cells(sourceRow, columnV, sourceRow, columnW)
            '    range.Merge = True
            '    range.Value = valueSource
            '    'range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            '    'range.Style.VerticalAlignment = ExcelVerticalAlignment.Center
            'End Using
        End If

        If targetColumn = "AD" Then
            valueColumn1 = If(IsNumeric(wsSource.Cells(sourceRow, columnAD).Value), CDbl(wsSource.Cells(sourceRow, columnAD).Value), 0)
            valueColumn2 = If(IsNumeric(wsSource.Cells(sourceRow, columnAE).Value), CDbl(wsSource.Cells(sourceRow, columnAE).Value), 0)
            valueSource = valueColumn1 + valueColumn2

            ' 合并 AD 列和 AE 列的单元格，并居中显示
            'Using range As ExcelRange = wsSource.Cells(sourceRow, columnAD, sourceRow, columnAE)
            '    range.Merge = True
            '    range.Value = valueSource
            '    'range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            '    'range.Style.VerticalAlignment = ExcelVerticalAlignment.Center
            'End Using
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
    Function ShouldSkipCell(cell As ExcelRangeBase, targetColumn As String, wsSource As ExcelWorksheet) As Boolean
        ' 检查是否是需要跳过的公司
        If cell.Value = "国网内蒙古东部电力有限公司经济技术研究院" Or
           cell.Value = "国网内蒙古东部电力有限公司内蒙古超特高压分公司" Or
           cell.Value = "内蒙古新正产业管理有限公司" Then
            ShouldSkipCell = True ' 跳过这些公司
            Exit Function
        End If

        ' 检查是否是需要特殊处理的公司
        If cell.Value = "国网内蒙古东部电力有限公司赤峰供电公司" Or
           cell.Value = "国网内蒙古东部电力有限公司通辽供电公司" Or
           cell.Value = "国网内蒙古东部电力有限公司兴安供电公司" Or
           cell.Value = "国网内蒙古东部电力有限公司呼伦贝尔供电公司" Then
            ' 如果是C列，且是这些公司，则不跳过
            If targetColumn = "C" Then ' 比较列名
                ShouldSkipCell = False ' C列不跳过
            Else
                ShouldSkipCell = True ' 非C列跳过
            End If
            Exit Function
        End If

        ' 获取目标列的单元格值
        Dim cellValue As String
        Dim columnIndex As Integer = ColumnLetterToNumber(targetColumn)
        cellValue = wsSource.Cells(cell.Start.Row, columnIndex).Value

        ' 默认返回是否为空或仅包含空格
        If Len(cellValue) = 0 Then
            ShouldSkipCell = True ' 目标列为空或者空格也跳过
        Else
            ShouldSkipCell = False ' 否则不跳过
        End If
    End Function
    ' 处理N602sheet anhuili
    Sub ProcessN602(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet)
        Dim cleanedValueA As String
        Dim returnedArray() As String
        Dim result As String ' 声明结果变量
        Dim SecondValue As String ' 声明第二个值变量
        Dim splitArray() As String ' 声明分割数组
        ' 获取使用区域
        Dim sourceEndRow As Integer = Math.Max(11, If(wsSource.Dimension?.End.Row, 11)) ' 确保 >= 11
        Dim targetEndRow As Integer = Math.Max(11, If(wsTarget.Dimension?.End.Row, 11)) ' 确保 >= 11
        Dim rangeSource As ExcelRange = wsSource.Cells($"C11:C{sourceEndRow}")
        Dim rangeTarget As ExcelRange = wsTarget.Cells($"C11:C{targetEndRow}")

        ' 调用函数并将返回的数组存储在Variant变量中
        ' 假设 CreateStringArrayForN602 是一个已经定义并返回数组的函数
        returnedArray = CreateStringArrayForN602()

        ' 遍历 rngN602 中的每个单元格
        For Each currentCell In rangeSource
            ' 为每个cell设置对应的AD单元格
            Dim cellC As ExcelRange = wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("C"))
            Dim cellAD As ExcelRange = wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("AD"))
            Dim cellAE As ExcelRange = wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("AE"))
            Dim cellA As ExcelRange = wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("A"))
            Dim cellValueTemp As String = cellA.Text

            '2.处理602表2栏和601表2栏一致
            For j As Integer = 1 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = wsTarget.Cells(j, ColumnLetterToNumber("A")).Value
                If cellValueTemp = targetCellValue Then
                    Dim targetBValue As Double = GetCellValueOrDefault(wsTarget.Cells(j, ColumnLetterToNumber("C")))
                    If cellC.Value = targetBValue.ToString And cellC.Value = cellAD.Value Then
                        ' 如果相等，两个单元格都标绿
                        SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber("C"), Color.Green) ' 一致标绿
                    Else
                        ' 如果不相等，两个单元格都标红
                        SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber("C"), Color.Red) ' 不一致标红
                        wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("C")).AddComment("N601 C列核对的数据为" & targetBValue & vbCrLf &
                          "N602 AD列核对的数据为" & cellAD.Value)
                    End If
                End If
            Next j

            ' 处理AD列如果单元格包含'['符号，则跳过本次循环，地市公司本部和旗县公司校验就行
            If Not cellValueTemp.Contains("【") Then
                Continue For
            End If
            cleanedValueA = RemoveBrackets(cellA.Text)
            ' 遍历 returnedArray 寻找匹配项
            For i = LBound(returnedArray) To UBound(returnedArray)
                If InStr(returnedArray(i), cleanedValueA) > 0 Then
                    result = returnedArray(i)
                    splitArray = Split(result, "=")
                    If UBound(splitArray) >= 1 Then
                        SecondValue = splitArray(1)
                        If cellAE.Value = SecondValue Then
                            SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber("AE"), Color.Green) ' 一致标绿
                        Else
                            SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber("AE"), Color.Red) ' 不一致标红
                            wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber("AE")).AddComment("核对的数据为" & SecondValue)
                        End If
                    End If
                    Exit For ' 如果找到匹配项，则退出循环
                End If
            Next i
        Next currentCell
    End Sub

    Public Sub ProcessN606CK(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, targetColumn As String)
        Dim rangeSource As ExcelRange
        Dim rangeTarget As ExcelRange
        Dim valueSource As Double

        ' 获取源工作表的 A 列范围（从第11行开始到最后一行）606
        ' 获取使用区域
        Dim sourceEndRow As Integer = Math.Max(11, If(wsSource.Dimension?.End.Row, 11)) ' 确保 >= 11
        rangeSource = wsSource.Cells($"A11:A{sourceEndRow}")

        ' 获取目标工作表的 A 列范围（从第4行开始到最后一行）N607-1
        Dim targetEndRow As Integer = Math.Max(11, If(wsTarget.Dimension?.End.Row, 11)) ' 确保 >= 11
        rangeTarget = wsTarget.Cells($"A11:A{targetEndRow}")

        ' 遍历源工作表 A 列
        For Each currentCell In rangeSource
            ' 清理当前源单元格 A 列的值，去除空格并去掉括号
            'cleanedValue = RemoveBrackets(currentCell.Value)
            ' 查找匹配值 

            ' 遍历目标工作表N607-1的 A 列，查找所有匹配的项
            For i As Integer = 11 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = wsTarget.Cells(i, ColumnLetterToNumber("A")).Value
                If currentCell.Value = targetCellValue Then
                    Dim targetBValue As Double = GetCellValueOrDefault(wsTarget.Cells(i, ColumnLetterToNumber("B")))
                    valueSource = GetCellValueOrDefault(wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber(targetColumn)))
                    If valueSource = targetBValue Then
                        ' 如果相等，两个单元格都标绿
                        SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber(targetColumn), Color.Green) ' 一致标绿
                    Else
                        ' 如果不相等，两个单元格都标红
                        SetCellColor(wsSource, currentCell.Start.Row, ColumnLetterToNumber(targetColumn), Color.Red) ' 不一致标红
                        wsSource.Cells(currentCell.Start.Row, ColumnLetterToNumber(targetColumn)).AddComment("N607-1核对的数据为" & targetBValue)
                    End If
                End If
            Next i
        Next currentCell
    End Sub
    Sub ProcessN606EF(wsSource As ExcelWorksheet, wsTarget As ExcelWorksheet, wsTargetKeMu As ExcelWorksheet, targetColumn As String)
        Dim cleanedValue As String
        Dim valueSource As Double
        Dim calculatedValue221113 As Double
        Dim calculatedValue221114 As Double
        Dim matchRowSum As Double
        Dim Value606 As Double

        ' 获取源和目标范围 ' 获取使用区域
        Dim sourceEndRow As Integer = Math.Max(11, If(wsSource.Dimension?.End.Row, 11)) ' 确保 >= 11
        Dim rangeSource As ExcelRange = wsSource.Cells($"A11:A{sourceEndRow}")  '607-3 607-4

        ' 遍历源工作表，遍历607-3或 607-4表，查出对应的1栏的值，去科目表查出对应的221113，221114
        For Each cell In rangeSource
            matchRowSum = 0 '初始化
            Value606 = 0
            '607-3 A 列的值
            cleanedValue = cell.Text
            '1栏的值
            valueSource = GetCellValueOrDefault(wsSource.Cells(cell.Start.Row, ColumnLetterToNumber("B")))
            For k As Integer = 4 To wsTargetKeMu.Dimension.End.Row
                Dim targetCellValue As String = wsTargetKeMu.Cells(k, ColumnLetterToNumber("A")).Value
                If cleanedValue.Contains(targetCellValue) Then '公司名
                    If wsTargetKeMu.Cells(k, ColumnLetterToNumber("B")).Value = "221113" Then
                        calculatedValue221113 = GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("E"))) + GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("F"))) - GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("I")))
                        matchRowSum = matchRowSum + calculatedValue221113
                    End If
                    If wsTargetKeMu.Cells(k, ColumnLetterToNumber("B")).Value = "221114" Then
                        calculatedValue221114 = GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("E"))) + GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("F"))) - GetCellValueOrDefault(wsTargetKeMu.Cells(k, ColumnLetterToNumber("I")))
                        matchRowSum = matchRowSum + calculatedValue221114
                    End If
                End If
            Next k
            '606栏的值
            For m = 11 To wsTarget.Dimension.End.Row
                Dim targetCellValue As String = wsTarget.Cells(m, ColumnLetterToNumber("A")).Value
                If cleanedValue = targetCellValue Then '公司名
                    Value606 = GetCellValueOrDefault(wsTarget.Cells(m, ColumnLetterToNumber(targetColumn)))
                    ' 处理V，W  AD，AE列
                    If targetColumn = "V" Then
                        Value606 = Value606 + GetCellValueOrDefault(wsTarget.Cells(m, ColumnLetterToNumber("W")))
                    End If
                    ' 处理V，W  AD，AE列
                    If targetColumn = "AD" Then
                        Value606 = Value606 + GetCellValueOrDefault(wsTarget.Cells(m, ColumnLetterToNumber("AE")))
                    End If
                    If Value606 <> 0 Then
                        If Math.Abs(Value606 - valueSource - matchRowSum) < 3 Then
                            SetCellColor(wsTarget, m, ColumnLetterToNumber(targetColumn), Color.Green) ' 一致标绿
                            If targetColumn = "AD" Then
                                SetCellColor(wsTarget, m, ColumnLetterToNumber("AE"), Color.Green) ' 一致标绿
                            End If
                            If targetColumn = "V" Then
                                SetCellColor(wsTarget, m, ColumnLetterToNumber("W"), Color.Green) ' 一致标绿
                            End If
                        Else
                            SetCellColor(wsTarget, m, ColumnLetterToNumber(targetColumn), Color.Red) ' 不一致标红
                            wsTarget.Cells(m, ColumnLetterToNumber(targetColumn)).AddComment("核对的数据为" & Math.Round((valueSource - matchRowSum), 2))
                            If targetColumn = "AD" Then
                                SetCellColor(wsTarget, m, ColumnLetterToNumber("AE"), Color.Red) ' 不一致标红
                            End If
                            If targetColumn = "V" Then
                                SetCellColor(wsTarget, m, ColumnLetterToNumber("W"), Color.Red) ' 不一致标红
                            End If
                        End If
                    End If
                End If
            Next m
        Next cell
    End Sub
    Public Function CreateStringArrayForN602() As String()
        ' 声明一个字符串数组，大小为56
        Dim stringArray(55) As String
        stringArray(0) = "国网内蒙古东部电力有限公司红山区供电分公司=2"
        stringArray(1) = "国网内蒙古东部电力有限公司本部=5"
        stringArray(2) = "国网内蒙古东部电力有限公司数字化事业部=3"
        stringArray(3) = "国网内蒙古东部电力有限公司物资事业部=3"
        stringArray(4) = "国网内蒙古东部电力有限公司经济技术研究院=3"
        stringArray(5) = "国网内蒙古东部电力有限公司内蒙古超特高压分公司=3"
        stringArray(6) = "国网内蒙古东部电力有限公司电力科学研究院=3"
        stringArray(7) = "国网内蒙古东部电力有限公司综合服务分公司=3"
        stringArray(8) = "国网内蒙古东部电力有限公司建设分公司=3"
        stringArray(9) = "国网内蒙古东部电力有限公司供电服务监管与支持中心=3"
        stringArray(10) = "国网内蒙古东部电力有限公司呼伦贝尔供电公司=3"
        stringArray(11) = "国网内蒙古东部电力有限公司呼伦贝尔供电公司本部=3"
        stringArray(12) = "国网内蒙古东部电力有限公司满洲里市供电分公司=2"
        stringArray(13) = "国网内蒙古东部电力有限公司根河市供电分公司=2"
        stringArray(14) = "国网内蒙古东部电力有限公司扎兰屯市供电分公司=2"
        stringArray(15) = "国网内蒙古东部电力有限公司牙克石市供电分公司=2"
        stringArray(16) = "国网内蒙古东部电力有限公司阿荣旗供电分公司=2"
        stringArray(17) = "国网内蒙古东部电力有限公司莫力达瓦达斡尔族自治旗供电分公司=2"
        stringArray(18) = "国网内蒙古东部电力有限公司鄂伦春自治旗供电分公司=2"
        stringArray(19) = "国网内蒙古东部电力有限公司新巴尔虎左旗供电分公司=2"
        stringArray(20) = "国网内蒙古东部电力有限公司新巴尔虎右旗供电分公司=2"
        stringArray(21) = "国网内蒙古东部电力有限公司陈巴尔虎旗供电分公司=2"
        stringArray(22) = "国网内蒙古东部电力有限公司额尔古纳市供电分公司=2"
        stringArray(23) = "国网内蒙古东部电力有限公司鄂温克族自治旗供电分公司=2"
        stringArray(24) = "国网内蒙古东部电力有限公司兴安供电公司=3"
        stringArray(25) = "国网内蒙古东部电力有限公司兴安供电公司本部=3"
        stringArray(26) = "国网内蒙古东部电力有限公司阿尔山市供电分公司=2"
        stringArray(27) = "国网内蒙古东部电力有限公司科右前旗供电分公司=2"
        stringArray(28) = "国网内蒙古东部电力有限公司突泉县供电分公司=2"
        stringArray(29) = "国网内蒙古东部电力有限公司科右中旗供电分公司=2"
        stringArray(30) = "国网内蒙古东部电力有限公司扎赉特旗供电分公司=2"
        stringArray(31) = "国网内蒙古东部电力有限公司乌兰浩特市供电分公司=2"
        stringArray(32) = "国网内蒙古东部电力有限公司通辽供电公司=3"
        stringArray(33) = "国网内蒙古东部电力有限公司通辽供电公司本部=3"
        stringArray(34) = "国网内蒙古东部电力有限公司库伦旗供电分公司=2"
        stringArray(35) = "国网内蒙古东部电力有限公司奈曼旗供电分公司=2"
        stringArray(36) = "国网内蒙古东部电力有限公司开鲁县供电分公司=2"
        stringArray(37) = "国网内蒙古东部电力有限公司科左后旗供电分公司=2"
        stringArray(38) = "国网内蒙古东部电力有限公司扎鲁特旗供电分公司=2"
        stringArray(39) = "国网内蒙古东部电力有限公司科左中旗供电分公司=2"
        stringArray(40) = "国网内蒙古东部电力有限公司科尔沁区供电分公司=2"
        stringArray(41) = "国网内蒙古东部电力有限公司新城区供电分公司=2"
        stringArray(42) = "国网内蒙古东部电力有限公司霍林郭勒市供电分公司=2"
        stringArray(43) = "国网内蒙古东部电力有限公司赤峰供电公司=2"
        stringArray(44) = "国网内蒙古东部电力有限公司元宝山区供电分公司=2"
        stringArray(45) = "国网内蒙古东部电力有限公司赤峰供电公司本部=3"
        stringArray(46) = "国网内蒙古东部电力有限公司阿鲁科尔沁旗供电分公司=2"
        stringArray(47) = "国网内蒙古东部电力有限公司巴林左旗供电分公司=2"
        stringArray(48) = "国网内蒙古东部电力有限公司巴林右旗供电分公司=2"
        stringArray(49) = "国网内蒙古东部电力有限公司林西县供电分公司=2"
        stringArray(50) = "国网内蒙古东部电力有限公司克什克腾旗供电分公司=2"
        stringArray(51) = "国网内蒙古东部电力有限公司翁牛特旗供电分公司=2"
        stringArray(52) = "国网内蒙古东部电力有限公司敖汉旗供电分公司=2"
        stringArray(53) = "国网内蒙古东部电力有限公司宁城县供电分公司=2"
        stringArray(54) = "国网内蒙古东部电力有限公司喀喇沁旗供电分公司=2"
        stringArray(55) = "国网内蒙古东部电力有限公司松山区供电分公司=2"
        Return stringArray
    End Function
    Public Function GetCellValueOrDefault(cell As ExcelRange) As Double
        ' 检查单元格值是否为 Nothing 或者空
        If cell.Value Is Nothing OrElse String.IsNullOrEmpty(cell.Value.ToString()) Then
            Return 0
        End If

        ' 尝试将单元格值转换为 Double 类型
        Dim value As Double
        If Double.TryParse(cell.Value.ToString(), value) Then
            ' 如果成功转换，则返回转换后的值
            Return value
        Else
            ' 如果不能转换为 Double，则返回 0
            Return 0
        End If
    End Function
End Class
