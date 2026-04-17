Option Explicit

'=============================================================
' CheckSeq - 序列检查与年假小时更新
' 功能：比对 IT2001 与 IT2006 工作表，按 SID 和日期区间匹配，
'        更新 IT2006 的累计小时数（第21列），并将结果写回 IT2001 第26列
' 作者：（你的名字）
' 修订：2026-04-17
'=============================================================
Sub CheckSeq()

    ' --- 工作表引用 ---
    Dim ws01 As Worksheet
    Dim ws06 As Worksheet
    Set ws01 = Worksheets("IT2001")
    Set ws06 = Worksheets("IT2006")

    ' --- 行边界 ---
    Dim lastRow01 As Long
    Dim lastRow06 As Long
    lastRow01 = ws01.Cells(ws01.Rows.Count, 1).End(xlUp).Row
    lastRow06 = ws06.Cells(ws06.Rows.Count, 1).End(xlUp).Row

    ' --- IT2001 数据数组（一次性读入，避免逐行访问工作表）---
    '    列映射：A=1, G=7, H=8, R=18, Y=25, Z=26
    Dim data01 As Variant
    data01 = ws01.Range(ws01.Cells(7, 1), ws01.Cells(lastRow01, 26)).Value

    ' --- IT2006 数据数组（一次性读入）---
    '    列映射：A=1, P=16, Q=17, R=18, S=19, T=20, U=21
    Dim data06 As Variant
    data06 = ws06.Range(ws06.Cells(7, 1), ws06.Cells(lastRow06, 21)).Value

    ' --- 性能优化 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 主循环：遍历 IT2001 每一行 ---
    Dim i As Long
    Dim j As Long

    ' IT2001 数组偏移量（数组从1开始，对应工作表第7行）
    Const OFFSET01 As Long = 6   ' 工作表行号 = 数组索引 + 6

    ' IT2006 数组偏移量
    Const OFFSET06 As Long = 6   ' 工作表行号 = 数组索引 + 6

    For i = 1 To UBound(data01, 1)

        ' 跳过已处理行（第26列 = 数组第26列）
        If data01(i, 26) <> "" Then GoTo NextI

        ' 读取 IT2001 当前行字段
        Dim SID01   As String
        Dim SDate01 As Date
        Dim EDate01 As Date
        Dim Hr01    As Double
        Dim crow01  As Long

        SID01   = CStr(data01(i, 1))
        SDate01 = CDate(data01(i, 7))
        EDate01 = CDate(data01(i, 8))
        Hr01    = CDbl(data01(i, 18))
        crow01  = CLng(data01(i, 25))

        If crow01 = 0 Then GoTo NextI

        ' --- 内循环：在 IT2006 中查找匹配行 ---
        Dim j06Start As Long
        j06Start = crow01 - OFFSET06   ' 转换为数组索引

        If j06Start < 1 Then j06Start = 1

        For j = j06Start To UBound(data06, 1) - 2  ' 留2行余量给 j+1 / j+2

            Dim SID06   As String
            Dim SDate06 As Date
            Dim EDate06 As Date
            Dim Hr06    As Double
            Dim Ded06   As Double
            Dim Acc06   As Double

            SID06   = CStr(data06(j, 1))
            SDate06 = CDate(data06(j, 17))
            EDate06 = CDate(data06(j, 18))
            Hr06    = CDbl(data06(j, 16))
            Ded06   = CDbl(data06(j, 19))
            Acc06   = CDbl(data06(j, 21))

            ' 条件：SID 匹配 + SDate01 在区间内 + 未达上限
            If SID06 = SID01 _
               And SDate01 >= SDate06 _
               And SDate01 <= EDate06 _
               And Acc06 < Ded06 Then

                Dim newAcc  As Double
                Dim overflow As Double
                newAcc   = Hr01 + Acc06
                overflow = newAcc - Ded06

                If newAcc >= Ded06 Then
                    ' 超出当前区间上限，需溢出到下一区间
                    data06(j, 21) = Ded06       ' 当前行填满

                    Dim nextRow As Long
                    Dim label06 As String

                    If Round(CDbl(data06(j + 1, 19)), 2) >= Round(overflow, 2) Then
                        ' 溢出量 <= 下一行上限：写入 j+1
                        nextRow = j + 1
                    Else
                        ' 溢出量 > 下一行上限：写入 j+2
                        nextRow = j + 2
                    End If

                    data06(nextRow, 21) = overflow
                    label06 = CStr(data06(j, 20)) & " " & CStr(data06(nextRow, 20))
                    data01(i, 26) = label06

                Else
                    ' 未超上限，直接累加
                    data06(j, 21) = newAcc
                    data01(i, 26) = CStr(data06(j, 20))
                End If

                Exit For    ' 找到匹配，退出内循环
            End If

        Next j

        ' 进度提示（每1000行显示一次，避免频繁刷新）
        If i Mod 1000 = 0 Then
            Application.StatusBar = "处理中：" & i & " / " & UBound(data01, 1) _
                                     & "  (" & Format(i / UBound(data01, 1), "0%") & ")"
        End If

NextI:
    Next i

    ' --- 将修改后的数组一次性回写工作表 ---
    ws01.Range(ws01.Cells(7, 1), ws01.Cells(lastRow01, 26)).Value = data01
    ws06.Range(ws06.Cells(7, 1), ws06.Cells(lastRow06, 21)).Value = data06

    ' --- 恢复设置 ---
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "处理完成！", vbInformation, "CheckSeq"

End Sub
