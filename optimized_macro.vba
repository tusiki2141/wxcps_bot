Option Explicit

' =====================================================
' 优化后的检查序列宏
' 功能：在IT2001和IT2006工作表之间进行工时匹配计算
' 优化重点：
' 1. 消除不必要的选择操作
' 2. 使用数组批量处理数据
' 3. 改进变量命名和代码结构
' 4. 优化循环逻辑
' =====================================================

Sub CheckseqOptimized()
    
    ' 声明变量
    Dim startTime As Single
    Dim ws2001 As Worksheet, ws2006 As Worksheet
    Dim lastRow2001 As Long, lastRow2006 As Long
    Dim data2001() As Variant, data2006() As Variant
    Dim i As Long, j As Long, z As Long
    Dim currentRow2001 As Long, currentRow2006 As Long
    Dim matchFound As Boolean
    
    ' 记录开始时间
    startTime = Timer
    
    ' 设置工作表引用
    Set ws2001 = Worksheets("IT2001")
    Set ws2006 = Worksheets("IT2006")
    
    ' 关闭屏幕更新和自动计算以提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 获取数据范围
    lastRow2001 = ws2001.Cells(ws2001.Rows.Count, "A").End(xlUp).Row
    lastRow2006 = ws2006.Cells(ws2006.Rows.Count, "A").End(xlUp).Row
    
    ' 检查是否有足够的数据
    If lastRow2001 < 7 Or lastRow2006 < 7 Then
        MsgBox "数据不足，请检查工作表数据"
        Exit Sub
    End If
    
    ' 将数据加载到数组中进行批量处理
    data2001 = ws2001.Range("A6:Z" & lastRow2001).Value
    data2006 = ws2006.Range("A7:V" & lastRow2006).Value
    
    ' 主处理循环 - 遍历IT2001数据
    For i = 1 To UBound(data2001, 1)
        
        ' 检查是否需要处理（第26列为空）
        If IsEmpty(data2001(i, 26)) Or data2001(i, 26) = "" Then
            
            ' 获取当前行数据
            Dim employeeId As String
            Dim startDate As Long, endDate As Long
            Dim hours2001 As Double
            
            employeeId = data2001(i, 1)
            startDate = data2001(i, 7)
            endDate = data2001(i, 8)
            hours2001 = data2001(i, 19)
            currentRow2006 = data2001(i, 25) ' 上次匹配的位置
            
            ' 更新状态栏显示进度
            Application.StatusBar = "进度: " & i & " / " & UBound(data2001, 1) & ": " & Format(i / UBound(data2001, 1), "Percent")
            
            ' 如果当前行有上次匹配记录，从该位置开始搜索
            If currentRow2006 = 0 Then currentRow2006 = 1
            
            matchFound = False
            
            ' 在IT2006中搜索匹配项
            For j = currentRow2006 To UBound(data2006, 1)
                
                ' 检查匹配条件
                If data2006(j, 1) = employeeId And _
                   startDate >= data2006(j, 17) And _
                   startDate <= data2006(j, 18) Then
                    
                    Dim availableHours As Double, usedHours As Double
                    availableHours = data2006(j, 16)
                    usedHours = data2006(j, 21)
                    
                    ' 情况1：有足够工时
                    If availableHours - usedHours - hours2001 >= 0 Then
                        data2006(j, 21) = usedHours + hours2001
                        data2001(i, 26) = data2006(j, 20)
                        data2001(i, 25) = j ' 记录匹配位置
                        matchFound = True
                        Exit For
                        
                    ' 情况2：工时不足，需要溢出处理
                    ElseIf availableHours - usedHours > 0 Then
                        Dim overflowHours As Double
                        overflowHours = (availableHours - usedHours - hours2001) * -1
                        
                        ' 分配当前可用的工时
                        data2006(j, 21) = availableHours
                        data2001(i, 26) = data2006(j, 20) & " " & overflowHours
                        
                        ' 处理溢出到下一行
                        matchFound = HandleHourOverflow(data2006, data2001, i, j, employeeId, startDate, overflowHours)
                        Exit For
                    End If
                End If
            Next j
            
            ' 如果没有找到匹配项，记录状态
            If Not matchFound Then
                data2001(i, 26) = "未找到匹配项"
            End If
        End If
    Next i
    
    ' 将处理后的数据写回工作表
    ws2001.Range("A6:Z" & lastRow2001).Value = data2001
    ws2006.Range("A7:V" & lastRow2006).Value = data2006
    
    ' 恢复应用程序设置
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    ' 保存工作簿
    ActiveWorkbook.Save
    
    ' 显示处理时间
    Dim elapsedTime As Single
    elapsedTime = Timer - startTime
    MsgBox "处理完成！耗时: " & Format(elapsedTime / 60, "0.00") & " 分钟"
    
End Sub

' =====================================================
' 处理工时溢出到下一行的函数
' =====================================================
Private Function HandleHourOverflow(ByRef data2006() As Variant, ByRef data2001() As Variant, _
                                   ByVal row2001 As Long, ByVal startRow2006 As Long, _
                                   ByVal employeeId As String, ByVal startDate As Long, _
                                   ByVal overflowHours As Double) As Boolean
    
    Dim result As Boolean
    result = False
    
    ' 检查下一行是否可以接收溢出工时
    If startRow2006 + 1 <= UBound(data2006, 1) Then
        
        Dim nextRow As Long
        nextRow = startRow2006 + 1
        
        ' 检查下一行是否满足条件
        If data2006(nextRow, 1) = employeeId And _
           startDate >= data2006(nextRow, 17) And _
           startDate <= data2006(nextRow, 18) And _
           data2006(nextRow, 16) - overflowHours > 0 Then
            
            ' 分配工时到下一行
            data2006(nextRow, 21) = overflowHours
            data2001(row2001, 26) = data2001(row2001, 26) & " " & data2006(nextRow, 20) & " " & overflowHours
            data2001(row2001, 25) = nextRow
            result = True
            
        Else
            ' 无法分配，标记为溢出
            data2006(startRow2006, 22) = "溢出 " & overflowHours
            data2001(row2001, 26) = data2001(row2001, 26) & " 溢出 " & overflowHours
            result = False
        End If
    Else
        ' 没有下一行，标记为溢出
        data2006(startRow2006, 22) = "溢出 " & overflowHours
        data2001(row2001, 26) = data2001(row2001, 26) & " 溢出 " & overflowHours
        result = False
    End If
    
    HandleHourOverflow = result
    
End Function

' =====================================================
' 辅助函数：清除之前的处理结果
' =====================================================
Sub ClearPreviousResults()
    
    Dim ws2001 As Worksheet, ws2006 As Worksheet
    Set ws2001 = Worksheets("IT2001")
    Set ws2006 = Worksheets("IT2006")
    
    ' 清除IT2001的第26列（处理结果）
    With ws2001
        If .Cells(.Rows.Count, "Z").End(xlUp).Row >= 6 Then
            .Range("Z6:Z" & .Cells(.Rows.Count, "Z").End(xlUp).Row).ClearContents
        End If
    End With
    
    ' 清除IT2006的第21-22列（已用工时和溢出标记）
    With ws2006
        If .Cells(.Rows.Count, "U").End(xlUp).Row >= 7 Then
            .Range("U7:V" & .Cells(.Rows.Count, "U").End(xlUp).Row).ClearContents
        End If
    End With
    
    MsgBox "已清除之前的处理结果"
    
End Sub