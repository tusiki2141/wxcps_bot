Option Explicit

Sub CheckSeq()

    ' ===== 变量声明 =====
    ' 工作表对象
    Dim ws01 As Worksheet
    Dim ws06 As Worksheet
    
    ' 行号
    Dim lastRow01   As Long   ' IT2001 数据末行
    Dim lastRow06   As Long   ' IT2006 数据末行
    Dim rowIdx01    As Long   ' IT2001 当前行索引
    Dim rowIdx06    As Long   ' IT2006 当前行索引
    Dim startRow06  As Long   ' IT2006 匹配起始行（来自 IT2001 第25列）
    
    ' IT2001 字段
    Dim empID01     As String  ' 员工ID（第1列）
    Dim startDate01 As Date    ' 开始日期（第7列）
    Dim endDate01   As Date    ' 结束日期（第8列）
    Dim hours01     As Double  ' 工时（第18列）
    Dim checkMark01 As String  ' 校验标记（第26列）
    
    ' IT2006 字段
    Dim empID06     As String  ' 员工ID（第1列）
    Dim startDate06 As Date    ' 合同开始日期（第17列）
    Dim endDate06   As Date    ' 合同结束日期（第18列）
    Dim hours06     As Double  ' 工时（第16列）
    Dim deduction06 As Double  ' 扣款阈值（第19列）
    Dim accum06     As Double  ' 已累计金额（第21列）
    
    ' 中间计算变量
    Dim nextDeduction As Double  ' 下一行的扣款值（第19列）
    Dim remainHours   As Double  ' 剩余工时 = hours01 + accum06 - deduction06
    
    ' ===== 性能优化：关闭屏幕刷新与自动计算 =====
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False   ' 额外禁用事件，减少触发
    
    ' ===== 绑定工作表对象，避免重复查找 =====
    Set ws01 = ThisWorkbook.Worksheets("IT2001")
    Set ws06 = ThisWorkbook.Worksheets("IT2006")
    
    ' ===== 获取数据末行（不依赖 Select/Activate）=====
    lastRow06 = ws06.Cells(ws06.Rows.Count, 1).End(xlUp).Row
    lastRow01 = ws01.Cells(ws01.Rows.Count, 1).End(xlUp).Row
    
    ' ===== 性能优化：将 IT2006 关键列读入数组，减少单元格 I/O =====
    Dim arr06 As Variant
    ' 读取 IT2006 第1~21列，从第7行到末行
    arr06 = ws06.Range(ws06.Cells(7, 1), ws06.Cells(lastRow06, 21)).Value
    ' arr06(r, c) 对应实际行 = r + 6，列 = c（1-based）
    '   列1  = 员工ID
    '   列16 = 工时
    '   列17 = 合同开始日期
    '   列18 = 合同结束日期
    '   列19 = 扣款阈值
    '   列20 = 标记字段
    '   列21 = 已累计金额

    ' ===== 主循环：遍历 IT2001 =====
    For rowIdx01 = 7 To lastRow01
    
        checkMark01 = ws01.Cells(rowIdx01, 26).Value
        
        ' 仅处理未标记的行
        If checkMark01 = "" Then
        
            empID01     = ws01.Cells(rowIdx01, 1).Value
            startDate01 = ws01.Cells(rowIdx01, 7).Value
            endDate01   = ws01.Cells(rowIdx01, 8).Value
            hours01     = ws01.Cells(rowIdx01, 18).Value
            startRow06  = ws01.Cells(rowIdx01, 25).Value
            
            ' 仅当存在有效起始行时才匹配
            If startRow06 <> 0 Then
            
                ' 将 startRow06 转换为数组下标（数组第1行对应实际第7行）
                Dim arrStartIdx As Long
                arrStartIdx = startRow06 - 6   ' arr06 的起始下标
                
                For rowIdx06 = arrStartIdx To UBound(arr06, 1)
                
                    empID06     = arr06(rowIdx06, 1)
                    startDate06 = arr06(rowIdx06, 17)
                    endDate06   = arr06(rowIdx06, 18)
                    hours06     = arr06(rowIdx06, 16)
                    deduction06 = arr06(rowIdx06, 19)
                    accum06     = arr06(rowIdx06, 21)
                    
                    ' 匹配条件：员工ID相同 + 日期在合同范围内 + 累计未达扣款阈值
                    If empID06 = empID01 _
                        And startDate01 >= startDate06 _
                        And startDate01 <= endDate06 _
                        And accum06 < deduction06 Then
                        
                        remainHours = hours01 + accum06 - deduction06
                        
                        If hours01 + accum06 >= deduction06 Then
                            ' 工时超过扣款阈值，需要拆分到下一行
                            nextDeduction = arr06(rowIdx06 + 1, 19)
                            
                            If Round(nextDeduction, 2) >= Round(remainHours, 2) Then
                                ' 情况A：余量可以放入 j+1 行
                                ws06.Cells(rowIdx06 + 6, 21).Value = remainHours
                                ws06.Cells(rowIdx06 + 5, 21).Value = deduction06
                                ' 更新数组缓存，保持一致性
                                arr06(rowIdx06 + 1, 21) = remainHours
                                arr06(rowIdx06, 21)     = deduction06
                                ws01.Cells(rowIdx01, 26).Value = _
                                    ws06.Cells(rowIdx06 + 5, 20).Value & " " & _
                                    ws06.Cells(rowIdx06 + 6, 20).Value
                            Else
                                ' 情况B：余量放入 j+2 行
                                ws06.Cells(rowIdx06 + 7, 21).Value = remainHours
                                ws06.Cells(rowIdx06 + 5, 21).Value = deduction06
                                arr06(rowIdx06 + 2, 21) = remainHours
                                arr06(rowIdx06, 21)     = deduction06
                                ws01.Cells(rowIdx01, 26).Value = _
                                    ws06.Cells(rowIdx06 + 5, 20).Value & " " & _
                                    ws06.Cells(rowIdx06 + 7, 20).Value
                            End If
                            
                        Else
                            ' 情况C：工时不超过阈值，直接累加
                            ws06.Cells(rowIdx06 + 5, 21).Value = hours01 + accum06
                            arr06(rowIdx06, 21) = hours01 + accum06
                            ws01.Cells(rowIdx01, 26).Value = ws06.Cells(rowIdx06 + 5, 20).Value
                        End If
                        
                        Exit For  ' 找到匹配行后退出内层循环
                    End If
                    
                Next rowIdx06
            End If
        End If
    Next rowIdx01
    
    ' ===== 恢复 Excel 默认设置 =====
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "CheckSeq 执行完成！", vbInformation

End Sub
