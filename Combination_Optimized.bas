Option Explicit

' ============================================================
' 模块级变量声明
' 说明：这些变量在多个子程序之间共享，需在模块顶部统一声明
' ============================================================

' 当前工作簿与控制工作表
Private wbThis As Workbook
Private wsCtrl As Worksheet

' 控制表数据数组（存储每步操作的所有参数）
Private actionSeq As Variant

' 控制表行数（操作步骤数量）与最后使用列号
Private stepCount As Integer
Private ctrlLastRow As Integer

' 错误信息（供验证函数使用）
Private ctrlErr As String

' 已打开/关闭的文件计数
Private cntOpen As Integer
Private cntClose As Integer

' 控制工作表名称列表（用于跳过控制表，不参与数据操作）
Private ctrlSheetNames As Variant


' ============================================================
' 主入口：Combination_main
' 功能：禁用屏幕刷新和提示，读取控制表B1单元格中的会话列表，
'       依次执行每个会话，完成后还原设置并弹出完成提示
' ============================================================
Sub Combination_main()

    Dim startTime As Date
    Dim sessionInput As String
    Dim sessionList() As String
    Dim i As Integer

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    startTime = Now

    ' 初始化控制工作表名称白名单
    ctrlSheetNames = Array("Control", "Src", "File", "Cut file", "Email")

    ' 激活本工作簿并定位控制工作表
    ThisWorkbook.Activate
    Set wbThis = ThisWorkbook
    Worksheets("Control").Activate
    Set wsCtrl = wbThis.ActiveSheet

    ' 读取 B1 单元格中的会话起始单元格（支持"|"分隔多个会话）
    sessionInput = wsCtrl.Range("B1").Value
    sessionList = Split(sessionInput, "|")

    ' 逐个执行会话
    For i = 0 To UBound(sessionList)
        Call Run_action(Trim(sessionList(i)))
    Next i

    ' 还原界面状态
    wbThis.Activate
    wsCtrl.Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wsCtrl.Range("B1").Select

    MsgBox "Completed." & vbNewLine & startTime & vbNewLine & Now

End Sub


' ============================================================
' 执行单个会话的所有操作步骤
' 参数：sessionStart — 控制表中该会话的起始单元格地址（如 "B3"）
' ============================================================
Sub Run_action(ByVal sessionStart As String)

    Dim i As Integer
    Dim actionName As String
    Dim currentRow As Integer
    Dim actionKey As String

    ' 读取控制表数据到 actionSeq 数组
    Call Get_control(sessionStart)

    ' 验证控制表输入（可在此函数内扩充校验逻辑）
    Call Validate_Control(sessionStart)

    ' 获取操作组名称（起始单元格上方第2行）和起始数据行号
    actionName = wsCtrl.Range(sessionStart).Offset(-2, 0).Value
    currentRow = wsCtrl.Range(sessionStart).Offset(1, 0).Row

    ' 逐步执行操作
    For i = 1 To stepCount

        Application.ScreenUpdating = False
        Application.StatusBar = "Processing... " & actionName & _
                                " Step " & i & " - Row_" & currentRow & _
                                "_" & actionSeq(i, 1) & _
                                ".  Total " & FormatPercent(i / stepCount) & " Completed."
        currentRow = currentRow + 1

        ' 将操作类型统一转为大写并去除空格，方便比较
        actionKey = UCase(Trim(actionSeq(i, 1)))

        ' ---- 使用 Select Case 替代大量 ElseIf，提升可读性与性能 ----
        Select Case actionKey

            Case "CLEAR_DATA"
                ' 参数: (工作簿, 工作表, 表头行, 起始数据行)
                Call Clear_Data(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 7))

            Case "COPY_FORMULA"
                ' 参数: (工作簿, 工作表, 指示行, 表头行, 公式行, 起始行)
                Call Copy_Formula(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 4), actionSeq(i, 6), actionSeq(i, 5), actionSeq(i, 7))

            Case "SAVE_FILE"
                ' 参数: (输出路径, 输出文件名)
                Call Save_File(actionSeq(i, 18), actionSeq(i, 19))

            Case "APPEND_ALL"
                Call Append_All(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 4), actionSeq(i, 6), actionSeq(i, 10), _
                                actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 16), actionSeq(i, 17))

            Case "APPEND_ALL_NOTCLOSEFILE"
                Call Append_All_notCloseFile(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 4), actionSeq(i, 6), actionSeq(i, 10), _
                                             actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 16), actionSeq(i, 17))

            Case "APPEND_BY_COL_NAME"
                Call Append_by_Col_Name(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 10), _
                                        actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 16), actionSeq(i, 17))

            Case "APPEND_BY_COL_NAME_NOTCLOSEFILE"
                Call Append_by_Col_Name_notCloseFile(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 10), _
                                                     actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 16), actionSeq(i, 17))

            Case "APPEND_IN_SAME_LINE"
                Call Append_in_Same_Line(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 4), actionSeq(i, 6), actionSeq(i, 7), _
                                         actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 16), actionSeq(i, 17))

            Case "REFRESH_PIVOT"
                Call Refresh_Pivot(actionSeq(i, 2), actionSeq(i, 3))

            Case "FILTER"
                Call Filter(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 9))

            Case "FILTER_PERIOD"
                Call Filter_Period(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 9))

            Case "SORTING"
                Call Sorting(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 8))

            Case "UNFILTER"
                Call Unfilter(actionSeq(i, 2), actionSeq(i, 3))

            Case "ADD_TEXT"
                Call add_text(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 7), actionSeq(i, 10))

            Case "CHANGE_TEXT"
                Call Change_text(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 7), actionSeq(i, 10))

            Case "DELETE_COL_ROW"
                Call Delete_Col_Row(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 11))

            Case "LIST_FILE"
                Call List_file(actionSeq(i, 12), actionSeq(i, 13), "FILE")

            Case "LIST_FILES"
                Call List_files(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 4), actionSeq(i, 15))

            Case "LIST_SUBFOLDER"
                Call List_subfolder(actionSeq(i, 12), "File")

            Case "LIST_ALL_FILES_SUBFOLDERS"
                Call List_all_files_subfolders(actionSeq(i, 12), "File")

            Case "COPY_ALL_FILES_SUBFOLDERS"
                Call Copy_all_files_subfolders(actionSeq(i, 12), actionSeq(i, 18))

            Case "RENAME_FILE"
                Call Rename_file(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 19))

            Case "COPY_SHEET"
                Call Copy_Sheet(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 14), actionSeq(i, 18), actionSeq(i, 19))

            Case "DELETE_SHEET"
                Call Delete_sheet(actionSeq(i, 2), actionSeq(i, 11))

            Case "CUT_FILE"
                Call Cut_File

            Case "PASTE_SHEET_AS_VALUE"
                Call Paste_sheet_as_value(actionSeq(i, 2), actionSeq(i, 3))

            Case "PASTE_CELL_AS_VALUE"
                Call Paste_cell_as_value(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 16), actionSeq(i, 17))

            Case "SEND_EMAIL"
                Call Send_Email(actionSeq(i, 6))

            Case "DELETE_FOLDER"
                Call Delete_folder(actionSeq(i, 12))

            Case "DELETE_FILE"
                Call Delete_file(actionSeq(i, 12), actionSeq(i, 13))

            Case "EXPAND_GROUP"
                Call Expand_group(actionSeq(i, 2), actionSeq(i, 3))

            Case "COLLAPSE_GROUP"
                Call Collapse_group(actionSeq(i, 2), actionSeq(i, 3))

            Case "PROTECT_SHEET"
                Call Protect_sheet(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 10))

            Case "UNPROTECT_SHEET"
                Call Unprotect_sheet(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 10))

            Case "PROTECT_WORKBOOK"
                Call Protect_workbook(actionSeq(i, 2), actionSeq(i, 10))

            Case "UNPROTECT_WORKBOOK"
                Call Unprotect_workbook(actionSeq(i, 2), actionSeq(i, 10))

            Case "TURN_ON_AUTO_CAL"
                Call Turn_on_auto_cal

            Case "TURN_OFF_AUTO_CAL"
                Call Turn_off_auto_cal

            Case "REMOVE_DUPLICATE"
                Call Remove_Duplicate(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 6), actionSeq(i, 9))

            Case "LIST_WORKSHEET"
                Call List_Worksheet(actionSeq(i, 12), actionSeq(i, 13))

            Case "PIVOT_FILTER"
                Call Pivot_Filter(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 9))

            Case "DATA_REFRESHALL"
                Call Data_RefreshAll(actionSeq(i, 2), actionSeq(i, 3))

            Case "COPY_SHEETDATA"
                Call Copy_sheetdata(actionSeq(i, 2), actionSeq(i, 3), actionSeq(i, 12), actionSeq(i, 13), _
                                    actionSeq(i, 14), actionSeq(i, 15), actionSeq(i, 17))

            Case Else
                ' 处理包含关键字的操作类型（OPEN_FILE / CLOSE_FILE / MOVE_FILE / COPY_FILE）
                ' 注：这些操作名称包含变体后缀（如 OPEN_FILE_EDIT），用 InStr 匹配更灵活
                If InStr(1, actionKey, "OPEN_FILE", vbTextCompare) > 0 Then
                    If InStr(1, actionKey, "EDIT", vbTextCompare) > 0 Then
                        Call Open_File(actionSeq(i, 12), actionSeq(i, 13), "EDIT")
                    Else
                        Call Open_File(actionSeq(i, 12), actionSeq(i, 13), "READ")
                    End If

                ElseIf InStr(1, actionKey, "CLOSE_FILE", vbTextCompare) > 0 Then
                    If InStr(1, actionKey, "SAVE AS", vbTextCompare) > 0 Then
                        Call Close_File(actionSeq(i, 18), actionSeq(i, 19), "SAVE AS")
                    ElseIf InStr(1, actionKey, "SAVE", vbTextCompare) > 0 Then
                        Call Close_File(actionSeq(i, 18), actionSeq(i, 19), "SAVE")
                    Else
                        Call Close_File(actionSeq(i, 18), actionSeq(i, 19), "UNCHANGE")
                    End If

                ElseIf InStr(1, actionKey, "MOVE_FILE", vbTextCompare) > 0 Then
                    If InStr(1, actionKey, "OVERWRITE", vbTextCompare) > 0 Then
                        Call Move_file(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 18), "OVERWRITE")
                    Else
                        Call Move_file(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 18), "")
                    End If

                ElseIf InStr(1, actionKey, "COPY_FILE", vbTextCompare) > 0 Then
                    If InStr(1, actionKey, "OVERWRITE", vbTextCompare) > 0 Then
                        Call Copy_file(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 18), "OVERWRITE")
                    Else
                        Call Copy_file(actionSeq(i, 12), actionSeq(i, 13), actionSeq(i, 18), "")
                    End If
                End If

        End Select

        Application.StatusBar = "Processing... Step " & i & " - " & actionSeq(i, 1) & _
                                 ".  Total " & FormatPercent(i / stepCount) & " Completed."

    Next i

    ' 清除状态栏并还原屏幕刷新
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub


' ============================================================
' 读取控制表数据到模块级数组 actionSeq
' 参数：ctrlStart — 该会话在控制表中的起始单元格地址
' 优化：改用 Range.Value 批量赋值，避免逐格 Select 带来的开销
' ============================================================
Sub Get_control(ByVal ctrlStart As String)

    Dim startCell As Range
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    wsCtrl.Activate
    Set startCell = wsCtrl.Range(ctrlStart)

    ' 计算步骤数量和最后使用行列
    lastRow = startCell.End(xlDown).Row
    stepCount = lastRow - startCell.Row
    ctrlLastRow = lastRow
    lastCol = startCell.End(xlToRight).Column

    ' 【性能优化】一次性将区域数据读入数组，避免循环中反复读取单元格
    ' 注意：从起始单元格下方第1行开始，共 stepCount 行、lastCol 列
    ReDim actionSeq(1 To stepCount, 1 To lastCol)

    Dim rawData As Variant
    Set dataRange = startCell.Offset(1, 0).Resize(stepCount, lastCol)
    rawData = dataRange.Value   ' 批量读取，rawData 为二维数组

    ' 将批量读取结果映射到 actionSeq（保持从索引1开始的约定）
    Dim r As Integer, c As Integer
    For r = 1 To stepCount
        For c = 1 To lastCol
            actionSeq(r, c) = rawData(r, c)
        Next c
    Next r

End Sub


' ============================================================
' 验证控制表输入的必填字段
' 参数：ctrlStart — 该会话在控制表中的起始单元格地址
' （当前为空实现，可根据需要扩充校验规则）
' ============================================================
Sub Validate_Control(ByVal ctrlStart As String)
    ' 预留：可在此添加字段必填检查、格式验证等逻辑
End Sub


' ============================================================
' 清除工作表表头行以下的所有数据行
' 若工作表含数据透视表，则仅清除非透视区域的列
' 参数：wkbk       — 工作簿名称（空则使用当前工作簿）
'       wksheet    — 工作表名称（空则使用当前活动工作表）
'       headerRow  — 表头所在行号
'       startRow   — 数据起始行号（空则取表头行+1）
' ============================================================
Sub Clear_Data(ByVal wkbk As String, ByVal wksheet As String, _
               ByVal headerRow As Variant, ByVal startRow As Variant)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    ' 确定起始行与最后数据行
    If startRow = "" Then startRow = headerRow + 1
    Dim lastDataRow As Long
    lastDataRow = last_row_check(headerRow)
    If lastDataRow = headerRow Then lastDataRow = lastDataRow + 1
    If lastDataRow < startRow Then lastDataRow = startRow

    If Worksheets(wksheet).PivotTables.Count > 0 Then
        ' 含数据透视表：只清除非透视区域的列
        Dim lastDataCol As Long
        lastDataCol = last_col_check(headerRow)
        For i = 1 To lastDataCol
            Cells(headerRow, i).Select
            If cell_in_pivot() = False Then
                Range(Cells(startRow, i), Cells(lastDataRow, i)).Clear
            End If
        Next i
    Else
        ' 无透视表：直接删除整行
        Range(Rows(startRow), Rows(lastDataRow)).EntireRow.Delete
    End If

End Sub


' ============================================================
' 将公式行的公式向下复制到所有数据行
' 针对指示行中标注为 "Formula" 的列执行复制-粘贴公式再转值
' 参数：wkbk       — 工作簿名称
'       wksheet    — 工作表名称
'       indRow     — 指示行行号（标注"Formula"的行）
'       headerRow  — 表头行行号
'       formulaRow — 存放模板公式的行号
'       startRow   — 数据起始行号
' ============================================================
Sub Copy_Formula(ByVal wkbk As String, ByVal wksheet As String, _
                 ByVal indRow As Variant, ByVal headerRow As Variant, _
                 ByVal formulaRow As Variant, ByVal startRow As Variant)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastCol As Long
    Dim lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    ' 确定数据起始行（跳过隐藏行）
    If startRow = "" Then
        If headerRow = lastDataRow Then Exit Sub
        For i = headerRow + 1 To lastDataRow
            If Not Rows(i).Hidden Then
                startRow = i
                Exit For
            End If
        Next i
    End If

    ' 遍历所有列，找到标注"Formula"的列并复制公式
    Range("A" & indRow).Select
    For i = 1 To lastCol
        If Selection.MergeCells = False Then
            If UCase(Selection.Value) = "FORMULA" Then
                Cells(formulaRow, Selection.Column).Copy
                Range(Cells(startRow, Selection.Column), Cells(lastDataRow, Selection.Column)).PasteSpecial xlPasteFormulasAndNumberFormats
                Call Unfilter(wkbk, wksheet)
                ' 将公式结果转为静态值，避免后续数据变动影响结果
                With Range(Cells(startRow, Selection.Column), Cells(lastDataRow, Selection.Column))
                    .Copy
                    .PasteSpecial xlPasteValues
                End With
                Application.CutCopyMode = False
            End If
        End If
        Cells(indRow, i + 1).Select
    Next i

End Sub


' ============================================================
' 打开工作簿文件
' 参数：srcPath  — 文件所在路径
'       srcFile  — 文件名
'       openType — "READ"（只读）或 "EDIT"（可编辑）
' ============================================================
Sub Open_File(ByVal srcPath As String, ByVal srcFile As String, ByVal openType As String)

    ' 确保路径末尾有反斜杠
    If Right(srcPath, 1) <> "\" Then srcPath = srcPath & "\"

    Select Case openType
        Case "READ"
            Workbooks.Open Filename:=srcPath & srcFile, ReadOnly:=True, UpdateLinks:=False
        Case "EDIT"
            Workbooks.Open Filename:=srcPath & srcFile, UpdateLinks:=False
    End Select

End Sub


' ============================================================
' 关闭工作簿文件
' 参数：outputPath — 保存路径（空则使用当前工作簿路径）
'       outputFile — 文件名
'       closeType  — "SAVE AS"（另存为）/ "SAVE"（保存关闭）/ 其他（不保存关闭）
' ============================================================
Sub Close_File(ByVal outputPath As String, ByVal outputFile As String, ByVal closeType As String)

    If outputPath = "" Then outputPath = ThisWorkbook.Path & "\"

    If closeType = "SAVE AS" Then
        If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
        ActiveWorkbook.SaveAs Filename:=outputPath & outputFile
    End If

    If workbook_is_open(outputFile) Then
        If closeType = "SAVE" Then
            Workbooks(outputFile).Close SaveChanges:=True
        Else
            Workbooks(outputFile).Close SaveChanges:=False
        End If
    End If

End Sub


' ============================================================
' 保存工作簿文件
' 参数：outputPath — 保存路径（空则使用当前工作簿路径）
'       outputFile — 文件名（空则使用当前工作簿名）
' ============================================================
Sub Save_File(ByVal outputPath As String, ByVal outputFile As String)

    If outputPath = "" Then
        outputPath = ThisWorkbook.Path & "\"
    Else
        If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
    End If

    If outputFile = "" Then outputFile = ThisWorkbook.Name

    If workbook_is_open(outputFile) Then
        Workbooks(outputFile).Save
    Else
        Workbooks(ThisWorkbook.Name).SaveAs outputPath & outputFile
    End If

End Sub


' ============================================================
' 将源文件全量数据追加到目标工作表（不检查列名匹配）
' 追加后自动关闭源文件
' 参数说明请参考控制表列定义
' ============================================================
Sub Append_All(ByVal wkbk As String, ByVal wksheet As String, _
               ByVal indRow As Variant, ByVal headerRow As Variant, _
               ByVal addSrcText As String, _
               ByVal srcPath As String, ByVal srcFile As String, _
               ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
               ByVal srcDataRow As Variant, ByVal srcDataCol As Variant)

    If srcFile = "N/A" Then Exit Sub

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Call Unfilter(wkbk, wksheet)
    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    ' 定位指示行中标注 SOURCE_START 的列
    Dim i As Integer
    Range("A" & indRow).Select
    For i = 1 To lastCol
        If UCase(Selection.Value) = "SOURCE_START" Then
            Cells(lastDataRow + 1, Selection.Column).Select

            ' 打开源文件（已打开则跳过）
            If (srcPath & srcFile) <> "" Then
                If workbook_is_open(srcFile) = False Then Call Open_File(srcPath, srcFile, "READ")
            Else
                srcFile = ThisWorkbook.Name
            End If

            If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
            Workbooks(srcFile).Worksheets(srcSheet).Activate

            If srcDataCol = "" Then srcDataCol = "A"
            Dim srcLastCell As String
            Dim srcLastRow As Long
            srcLastCell = last_cell_check(srcHeaderRow)
            srcLastRow = last_row_check(srcHeaderRow)

            If srcLastRow >= srcDataRow Then
                Range(Cells(srcDataRow, srcDataCol), srcLastCell).Copy
                Workbooks(wkbk).Worksheets(wksheet).Activate
                ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
                Application.CutCopyMode = False

                ' 写入来源标注文本（控制表 J 列配置）
                Call WriteSourceText(addSrcText, lastDataRow, srcLastRow, srcDataRow)
            End If

            ' 关闭源文件
            If (srcPath & srcFile) <> "" And srcFile <> ThisWorkbook.Name Then
                Call Close_File(srcPath, srcFile, "UNCHANGE")
            End If

            Exit For
        End If
        Selection.Offset(0, 1).Select
    Next i

    Range("A" & indRow).Select

End Sub


' ============================================================
' 将源文件全量数据追加到目标工作表（不关闭源文件版本）
' ============================================================
Sub Append_All_notCloseFile(ByVal wkbk As String, ByVal wksheet As String, _
                             ByVal indRow As Variant, ByVal headerRow As Variant, _
                             ByVal addSrcText As String, _
                             ByVal srcPath As String, ByVal srcFile As String, _
                             ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
                             ByVal srcDataRow As Variant, ByVal srcDataCol As Variant)

    If srcFile = "N/A" Then Exit Sub

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Call Unfilter(wkbk, wksheet)
    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    Dim i As Integer
    Range("A" & indRow).Select
    For i = 1 To lastCol
        If UCase(Selection.Value) = "SOURCE_START" Then
            Cells(lastDataRow + 1, Selection.Column).Select

            If (srcPath & srcFile) <> "" Then
                If workbook_is_open(srcFile) = False Then Call Open_File(srcPath, srcFile, "READ")
            Else
                srcFile = ThisWorkbook.Name
            End If

            If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
            Workbooks(srcFile).Worksheets(srcSheet).Activate

            If srcDataCol = "" Then srcDataCol = "A"
            Dim srcLastCell As String
            Dim srcLastRow As Long
            srcLastCell = last_cell_check(srcHeaderRow)
            srcLastRow = last_row_check(srcHeaderRow)

            If srcLastRow >= srcDataRow Then
                Range(Cells(srcDataRow, srcDataCol), srcLastCell).Copy
                Workbooks(wkbk).Worksheets(wksheet).Activate
                ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
                Application.CutCopyMode = False
                Call WriteSourceText(addSrcText, lastDataRow, srcLastRow, srcDataRow)
            End If

            Exit For
        End If
        Selection.Offset(0, 1).Select
    Next i

    Range("A" & indRow).Select

End Sub


' ============================================================
' 按列名匹配，将源文件数据追加到目标工作表（追加后关闭源文件）
' ============================================================
Sub Append_by_Col_Name(ByVal wkbk As String, ByVal wksheet As String, _
                        ByVal headerRow As Variant, ByVal addSrcText As String, _
                        ByVal srcPath As String, ByVal srcFile As String, _
                        ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
                        ByVal srcDataRow As Variant, ByVal srcDataCol As Variant)

    If srcFile = "N/A" Then Exit Sub

    If (srcPath & srcFile) <> "" Then
        If workbook_is_open(srcFile) = False Then Call Open_File(srcPath, srcFile, "READ")
    Else
        srcFile = ThisWorkbook.Name
    End If

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
    Workbooks(srcFile).Worksheets(srcSheet).Activate

    If srcDataCol = "" Then srcDataCol = "A"
    Dim srcStartCol As Long, srcLastCol As Long, srcLastRow As Long
    srcStartCol = Range(srcDataCol & srcHeaderRow).Column
    srcLastCol = last_col_check(srcHeaderRow)
    srcLastRow = last_row_check(srcHeaderRow)

    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    Dim i As Integer, j As Integer
    If srcLastRow >= srcDataRow Then
        For i = 1 To lastCol
            For j = srcStartCol To srcLastCol
                Dim destHeader As String, srcHeader As String
                destHeader = Workbooks(wkbk).Worksheets(wksheet).Cells(headerRow, i).Value
                srcHeader = Workbooks(srcFile).Worksheets(srcSheet).Cells(srcHeaderRow, j).Value

                If Trim(destHeader) = Trim(srcHeader) Then
                    Workbooks(srcFile).Worksheets(srcSheet).Range( _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcDataRow, j), _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcLastRow, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(lastDataRow + 1, i).PasteSpecial xlPasteValuesAndNumberFormats

                    Call WriteSourceText(addSrcText, lastDataRow, srcLastRow, srcDataRow)
                    Exit For
                End If
            Next j
        Next i
    End If

    Range("A" & headerRow).Select
    Application.CutCopyMode = False

    If (srcPath & srcFile) <> "" And srcFile <> ThisWorkbook.Name Then
        Call Close_File(srcPath, srcFile, "UNCHANGE")
    End If

End Sub


' ============================================================
' 按列名匹配追加数据（不关闭源文件版本）
' ============================================================
Sub Append_by_Col_Name_notCloseFile(ByVal wkbk As String, ByVal wksheet As String, _
                                     ByVal headerRow As Variant, ByVal addSrcText As String, _
                                     ByVal srcPath As String, ByVal srcFile As String, _
                                     ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
                                     ByVal srcDataRow As Variant, ByVal srcDataCol As Variant)

    If srcFile = "N/A" Then Exit Sub

    If (srcPath & srcFile) <> "" Then
        If workbook_is_open(srcFile) = False Then Call Open_File(srcPath, srcFile, "READ")
    Else
        srcFile = ThisWorkbook.Name
    End If

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
    Workbooks(srcFile).Worksheets(srcSheet).Activate

    If srcDataCol = "" Then srcDataCol = "A"
    Dim srcStartCol As Long, srcLastCol As Long, srcLastRow As Long
    srcStartCol = Range(srcDataCol & srcHeaderRow).Column
    srcLastCol = last_col_check(srcHeaderRow)
    srcLastRow = last_row_check(srcHeaderRow)

    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    Dim i As Integer, j As Integer
    If srcLastRow >= srcDataRow Then
        For i = 1 To lastCol
            For j = srcStartCol To srcLastCol
                Dim destHeader As String, srcHeader As String
                destHeader = Workbooks(wkbk).Worksheets(wksheet).Cells(headerRow, i).Value
                srcHeader = Workbooks(srcFile).Worksheets(srcSheet).Cells(srcHeaderRow, j).Value

                If Trim(destHeader) = Trim(srcHeader) Then
                    Workbooks(srcFile).Worksheets(srcSheet).Range( _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcDataRow, j), _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcLastRow, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(lastDataRow + 1, i).PasteSpecial xlPasteValuesAndNumberFormats

                    Call WriteSourceText(addSrcText, lastDataRow, srcLastRow, srcDataRow)
                    Exit For
                End If
            Next j
        Next i
    End If

    Range("A" & headerRow).Select
    Application.CutCopyMode = False

End Sub


' ============================================================
' 私有辅助过程：写入来源标注文本
' 将 addSrcText 解析后填入目标工作表指定列
' 参数：addSrcText  — 格式为 "列名|值1,值2;列名2|值"
'       baseRow     — 目标工作表当前最后一行（追加前）
'       srcLastRow  — 源数据最后行
'       srcDataRow  — 源数据起始行
' ============================================================
Private Sub WriteSourceText(ByVal addSrcText As String, _
                             ByVal baseRow As Long, ByVal srcLastRow As Long, ByVal srcDataRow As Long)

    If addSrcText = "" Or addSrcText = "NA" Then Exit Sub

    Dim details() As String
    details = Split(addSrcText, ";")
    Dim k As Integer
    For k = 0 To UBound(details)
        Dim sepPos As Long
        sepPos = InStr(1, details(k), "|", vbTextCompare)
        Dim toCol As Long
        toCol = Range(Trim(Left(details(k), sepPos - 1)) & "1").Column
        Dim addValue As Variant
        addValue = Split(Trim(Mid(details(k), sepPos + 1)), ",")
        Range(Cells(baseRow + 1, toCol), Cells(baseRow + srcLastRow - srcDataRow + 1, toCol)).Value = addValue
    Next k

End Sub


' ============================================================
' 按行与源文件对应列名匹配，在同行中填入数据（不新增行）
' ============================================================
Sub Append_in_Same_Line(ByVal wkbk As String, ByVal wksheet As String, _
                         ByVal indRow As Variant, ByVal headerRow As Variant, _
                         ByVal startRow As Variant, _
                         ByVal srcPath As String, ByVal srcFile As String, _
                         ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
                         ByVal srcDataRow As Variant, ByVal srcDataCol As Variant)

    Application.DisplayAlerts = False

    If srcFile = "N/A" Then Exit Sub

    If (srcPath & srcFile) <> "" Then
        If workbook_is_open(srcFile) = False Then Call Open_File(srcPath, srcFile, "READ")
    Else
        srcFile = ThisWorkbook.Name
    End If

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
    Workbooks(srcFile).Worksheets(srcSheet).Activate
    Call Unfilter(srcFile, srcSheet)

    If srcDataCol = "" Then srcDataCol = "A"
    Dim srcStartCol As Long, srcLastCol As Long, srcLastRow As Long
    srcStartCol = Range(srcDataCol & srcHeaderRow).Column
    srcLastCol = last_col_check(srcHeaderRow)
    srcLastRow = last_row_check(srcHeaderRow)

    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    ' 找到 SOURCE_START / SOURCE_END 对应列
    Dim srcStartDataCol As Long, srcEndDataCol As Long
    Dim i As Integer
    For i = 1 To lastCol
        If UCase(Cells(indRow, i)) = "SOURCE_START" Then srcStartDataCol = i
        If UCase(Cells(indRow, i)) = "SOURCE_END" Then srcEndDataCol = i
    Next i

    If srcLastRow >= srcDataRow Then
        Dim j As Integer
        For i = srcStartDataCol To srcEndDataCol
            For j = srcStartCol To srcLastCol
                Dim destHeader As String, srcHeader As String
                destHeader = Workbooks(wkbk).Worksheets(wksheet).Cells(headerRow, i).Value
                srcHeader = Workbooks(srcFile).Worksheets(srcSheet).Cells(srcHeaderRow, j).Value

                If Trim(destHeader) = Trim(srcHeader) Then
                    Workbooks(srcFile).Worksheets(srcSheet).Range( _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcDataRow, j), _
                        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcLastRow, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(startRow, i).PasteSpecial xlPasteValuesAndNumberFormats
                    Exit For
                End If
            Next j
        Next i
    End If

    Range("A" & headerRow).Select
    Application.CutCopyMode = False

    If (srcPath & srcFile) <> "" And srcFile <> ThisWorkbook.Name Then
        Call Close_File(srcPath, srcFile, "UNCHANGE")
    End If

End Sub


' ============================================================
' 刷新工作簿中的数据透视表
' 参数：wkbk    — 工作簿名称
'       wksheet — 工作表名称（"ALL" 表示刷新所有非控制工作表的透视表）
' ============================================================
Sub Refresh_Pivot(ByVal wkbk As String, ByVal wksheet As String)

    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name

    If UCase(wksheet) = "ALL" Then
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                For Each PT In ws.PivotTables
                    PT.SourceData = check_pivot_range(PT.SourceData)
                    PT.RefreshTable
                Next PT
            End If
        Next ws
    Else
        For Each PT In Workbooks(wkbk).Worksheets(wksheet).PivotTables
            PT.RefreshTable
        Next PT
    End If

End Sub


' ============================================================
' 按用户指定列名和值进行精确筛选
' 参数：filter_by 格式：列字母|值1,值2;列字母2|值...
' ============================================================
Sub Filter(ByVal wkbk As String, ByVal wksheet As String, _
           ByVal headerRow As Variant, ByVal filterBy As String)

    Dim i As Integer, j As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    Dim filterDetails() As String
    filterDetails = Split(filterBy, ";")

    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False

    For i = 0 To UBound(filterDetails)
        Dim sepPos As Long
        sepPos = InStr(1, filterDetails(i), "|", vbTextCompare)
        Dim filterCol As Long
        filterCol = Range(Trim(Left(filterDetails(i), sepPos - 1)) & "1").Column
        Dim filterValues() As String
        filterValues = Split(Trim(Mid(filterDetails(i), sepPos + 1)), ",")

        ' 将 "(blank)" 替换为空字符串
        For j = 0 To UBound(filterValues)
            If Trim(filterValues(j)) = "(blank)" Then filterValues(j) = ""
        Next j

        ' 应用筛选
        If UBound(filterValues) = 0 And filterValues(0) = "<>" Then
            ActiveSheet.Rows(headerRow & ":" & lastDataRow).AutoFilter _
                Field:=filterCol, Criteria1:="<>", Operator:=xlFilterValues
        Else
            ActiveSheet.Rows(headerRow & ":" & lastDataRow).AutoFilter _
                Field:=filterCol, Criteria1:=filterValues, Operator:=xlFilterValues
        End If
    Next i

End Sub


' ============================================================
' 按日期或比较条件进行区间筛选
' ============================================================
Sub Filter_Period(ByVal wkbk As String, ByVal wksheet As String, _
                  ByVal headerRow As Variant, ByVal filterBy As String)

    Dim i As Integer, j As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastCol As Long, lastDataRow As Long
    lastCol = last_col_check(headerRow)
    lastDataRow = last_row_check(headerRow)

    Dim filterDetails() As String
    filterDetails = Split(filterBy, ";")

    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False

    For i = 0 To UBound(filterDetails)
        Dim sepPos As Long
        sepPos = InStr(1, filterDetails(i), "|", vbTextCompare)
        Dim filterCol As Long
        filterCol = Range(Trim(Left(filterDetails(i), sepPos - 1)) & "1").Column
        Dim filterValues() As String
        filterValues = Split(Trim(Mid(filterDetails(i), sepPos + 1)), ",")

        For j = 0 To UBound(filterValues)
            If Trim(filterValues(j)) = "(blank)" Then filterValues(j) = ""
        Next j

        ' 日期或比较运算符条件用 xlAnd，否则用列表筛选
        If IsDate(Right(filterValues(0), 10)) Or _
           Left(filterValues(0), 1) = ">" Or Left(filterValues(0), 1) = "<" Then
            If UBound(filterValues) = 0 Then
                ActiveSheet.Rows(headerRow & ":" & lastDataRow).AutoFilter _
                    Field:=filterCol, Criteria1:=filterValues(0), Operator:=xlAnd
            Else
                ActiveSheet.Rows(headerRow & ":" & lastDataRow).AutoFilter _
                    Field:=filterCol, Criteria1:=filterValues(0), Operator:=xlAnd, Criteria2:=filterValues(1)
            End If
        Else
            ActiveSheet.Rows(headerRow & ":" & lastDataRow).AutoFilter _
                Field:=filterCol, Criteria1:=filterValues, Operator:=xlFilterValues
        End If
    Next i

End Sub


' ============================================================
' 按指定列和排序方式对数据区域进行排序
' 参数：sortBy 格式：列字母|Ascending 或 列字母|Descending(自定义1,自定义2,...)
' ============================================================
Sub Sorting(ByVal wkbk As String, ByVal wksheet As String, _
            ByVal headerRow As Variant, ByVal sortBy As String)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastDataRow As Long, lastCell As String
    lastDataRow = last_row_check(headerRow)
    lastCell = last_cell_check(headerRow)

    ActiveSheet.Sort.SortFields.Clear
    Dim sortDetails() As String
    sortDetails = Split(sortBy, ";")

    Dim sort_custom_count As Integer
    sort_custom_count = 0

    For i = 0 To UBound(sortDetails)
        Dim sepPos As Long, customPos As Long
        sepPos = InStr(1, sortDetails(i), "|", vbTextCompare)
        customPos = InStr(1, sortDetails(i), "(", vbTextCompare)

        Dim sortCol As String, sortOrder As String
        sortCol = Trim(Left(sortDetails(i), sepPos - 1))

        If customPos = 0 Then
            ' 普通升降序
            sortOrder = Trim(Mid(sortDetails(i), sepPos + 1))
        Else
            ' 含自定义排序列表
            sortOrder = Trim(Mid(sortDetails(i), sepPos + 1, customPos - sepPos - 1))
            Dim sortCustomList() As String
            sortCustomList = Split(Mid(sortDetails(i), customPos + 1, Len(sortDetails(i)) - customPos - 1), ",")
            Application.AddCustomList sortCustomList
            sort_custom_count = Application.CustomListCount
        End If

        If UCase(sortOrder) = "ASCENDING" Then
            ActiveSheet.Sort.SortFields.Add _
                Key:=Range(sortCol & headerRow + 1 & ":" & sortCol & lastDataRow), _
                SortOn:=xlSortOnValues, Order:=xlAscending, _
                CustomOrder:=sort_custom_count, DataOption:=xlSortNormal
        ElseIf UCase(sortOrder) = "DESCENDING" Then
            ActiveSheet.Sort.SortFields.Add _
                Key:=Range(sortCol & headerRow + 1 & ":" & sortCol & lastDataRow), _
                SortOn:=xlSortOnValues, Order:=xlDescending, _
                CustomOrder:=sort_custom_count, DataOption:=xlSortNormal
        End If
    Next i

    ' 应用排序设置
    With ActiveSheet.Sort
        .SetRange Range("A" & headerRow & ":" & lastCell)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


' ============================================================
' 取消工作表筛选（清除筛选条件，显示所有数据行）
' 参数：wkbk    — 工作簿名称（"ALL" 表示处理所有非控制工作表）
'       wksheet — 工作表名称
' ============================================================
Sub Unfilter(ByVal wkbk As String, ByVal wksheet As String)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Activate
                Call ClearSheetFilter(ws)
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Call ClearSheetFilter(Workbooks(wkbk).Worksheets(wksheet))
    End If

End Sub


' ============================================================
' 私有辅助过程：清除单个工作表的筛选
' ============================================================
Private Sub ClearSheetFilter(ByVal ws As Worksheet)
    If ws.AutoFilterMode Then
        Dim rng As Range
        Set rng = ws.AutoFilter.Range
        If rng.Rows.Count > rng.SpecialCells(xlCellTypeVisible).Rows.Count Then
            ws.ShowAllData
        Else
            ws.AutoFilterMode = False
            rng.AutoFilter
        End If
    End If
End Sub


' ============================================================
' 向筛选后可见行的指定列写入固定文本
' 参数：text_to_add 格式：列字母|文本;列字母2|文本2...
' ============================================================
Sub add_text(ByVal wkbk As String, ByVal wksheet As String, _
             ByVal headerRow As Variant, ByVal startRow As Variant, _
             ByVal textToAdd As String)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastDataRow As Long
    lastDataRow = last_row_check(headerRow)
    If startRow = "" Then startRow = headerRow + 1
    If lastDataRow = headerRow Then Exit Sub

    Dim addDetails() As String
    addDetails = Split(textToAdd, ";")

    For i = 0 To UBound(addDetails)
        Dim sepPos As Long
        sepPos = InStr(1, addDetails(i), "|", vbTextCompare)
        Dim toAddCol As String, toAddText As String
        toAddCol = Trim(Left(addDetails(i), sepPos - 1))
        toAddText = Trim(Mid(addDetails(i), sepPos + 1))

        ' 仅对可见（未隐藏）行写入文本
        Dim cell As Range
        For Each cell In Range(toAddCol & CStr(startRow) & ":" & toAddCol & CStr(lastDataRow))
            If Not cell.Rows.Hidden Then
                cell.Value = toAddText
            End If
        Next cell
    Next i

End Sub


' ============================================================
' 修改指定单元格区域内可见行的值
' 参数：textToChange 格式：单元格地址范围|新文本;...
' ============================================================
Sub Change_text(ByVal wkbk As String, ByVal wksheet As String, _
                ByVal headerRow As Variant, ByVal startRow As Variant, _
                ByVal textToChange As String)

    Dim i As Integer

    Application.Calculation = xlAutomatic

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim addDetails() As String
    addDetails = Split(textToChange, ";")

    For i = 0 To UBound(addDetails)
        Dim sepPos As Long
        sepPos = InStr(1, addDetails(i), "|", vbTextCompare)
        Dim toAddRange As String, toAddText As String
        toAddRange = Trim(Left(addDetails(i), sepPos - 1))
        toAddText = Trim(Mid(addDetails(i), sepPos + 1))

        Dim cell As Range
        For Each cell In Range(toAddRange)
            If Not cell.Rows.Hidden Then
                cell.Value = toAddText
            End If
        Next cell
    Next i

End Sub


' ============================================================
' 删除指定行或列
' 参数：del_by 格式：ROW|行号1,行号2;COLUMN|列字母...
' ============================================================
Sub Delete_Col_Row(ByVal wkbk As String, ByVal wksheet As String, ByVal delBy As String)

    Dim i As Integer, j As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim delDetails() As String
    delDetails = Split(delBy, ";")

    For i = 0 To UBound(delDetails)
        Dim sepPos As Long
        sepPos = InStr(1, delDetails(i), "|", vbTextCompare)
        Dim delType As String
        delType = Trim(Left(delDetails(i), sepPos - 1))
        Dim delItems() As String
        delItems = Split(Trim(Mid(delDetails(i), sepPos + 1)), ",")

        ' 拼接选取地址字符串（例如 "3:3,5:5"）
        Dim delItemsStr As String
        delItemsStr = ""
        For j = 0 To UBound(delItems)
            Dim item As String
            item = Trim(delItems(j))
            If InStr(1, item, ":") = 0 Then item = item & ":" & item
            If delItemsStr = "" Then
                delItemsStr = item
            Else
                delItemsStr = delItemsStr & "," & item
            End If
        Next j

        If UCase(delType) = "ROW" Then
            Range(delItemsStr).EntireRow.Delete
        ElseIf UCase(delType) = "COLUMN" Then
            Range(delItemsStr).EntireColumn.Delete
        End If
    Next i

End Sub


' ============================================================
' 列出文件夹内指定类型的文件
' 参数：srcPath  — 文件夹路径
'       srcType  — 文件扩展名通配符（如 "*.xlsx"）
'       wksheet  — 写入文件列表的工作表名称
' ============================================================
Sub List_file(ByVal srcPath As String, ByVal srcType As String, ByVal wksheet As String)

    If Right(srcPath, 1) <> "\" Then srcPath = srcPath & "\"

    Dim fileName As String
    fileName = Dir(srcPath & srcType)

    Worksheets(wksheet).Activate
    Dim lastDataRow As Long
    lastDataRow = last_row_check(1)
    Range("A" & lastDataRow).Select

    Do While fileName <> ""
        ' 扩展名匹配检查（兼容 *xls = *xlsx 的情况）
        If Right(fileName, Len(srcType) - 1) = Right(srcType, Len(srcType) - 1) Or srcType = "*.*" Then
            Selection.Offset(1, 0).Select
            ActiveCell.Value = srcPath
            ActiveCell.Offset(0, 1).Value = fileName
        End If
        fileName = Dir
    Loop

End Sub


' ============================================================
' 列出文件夹内指定类型的文件，写入到指示行对应列
' ============================================================
Sub List_files(ByVal srcPath As String, ByVal srcType As String, _
               ByVal wksheet As String, ByVal indRow As Variant, ByVal headerRow As Variant)

    If Right(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
    If wksheet = "" Then wksheet = "File"

    Dim fileName As String
    fileName = Dir(srcPath & srcType)

    Worksheets(wksheet).Activate
    Dim lastDataRow As Long, lastCol As Long
    lastDataRow = last_row_check(1)
    lastCol = last_col_check(1)

    ' 找到指示行中标注 LIST_FILES 的列
    Dim i As Integer
    For i = 1 To lastCol
        If UCase(Cells(indRow, i).Value) = "LIST_FILES" Then
            Cells(lastDataRow, i).Select

            Do While fileName <> ""
                If UCase(Right(fileName, Len(srcType) - 1)) = UCase(Right(srcType, Len(srcType) - 1)) Or srcType = "*.*" Then
                    Selection.Offset(1, 0).Select
                    ActiveCell.Value = srcPath
                    ActiveCell.Offset(0, 1).Value = fileName
                End If
                fileName = Dir
            Loop

            Exit For
        End If
    Next i

End Sub


' ============================================================
' 列出文件夹内所有子文件夹名
' ============================================================
Sub List_subfolder(ByVal srcPath As String, ByVal wksheet As String)

    Dim fso As Object, folder As Object, subFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(srcPath)

    Worksheets(wksheet).Activate
    Dim lastDataRow As Long
    lastDataRow = last_row_check(1)
    Range("A" & lastDataRow).Select

    For Each subFolder In folder.SubFolders
        Selection.Offset(1, 0).Select
        ActiveCell.Value = srcPath
        ActiveCell.Offset(0, 1).Value = subFolder.Name
    Next subFolder

End Sub


' ============================================================
' 递归列出文件夹内所有子文件夹和文件
' ============================================================
Sub List_all_files_subfolders(ByVal srcPath As String, ByVal wksheet As String)

    Dim fso As Object, folder As Object, subFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(srcPath)

    Worksheets(wksheet).Activate
    Dim lastDataRow As Long
    lastDataRow = last_row_check(1)
    Range("A" & lastDataRow).Select

    For Each subFolder In folder.SubFolders
        Selection.Offset(1, 0).Select
        ActiveCell.Value = srcPath
        ActiveCell.Offset(0, 1).Value = subFolder.Name
        ' 递归处理子文件夹
        Call List_all_files_subfolders(srcPath & "\" & subFolder.Name, wksheet)
    Next subFolder

    Call List_file(srcPath, "*.*", wksheet)

End Sub


' ============================================================
' 移动文件到目标文件夹
' 参数：moveType — "OVERWRITE"（覆盖已存在文件）或 ""（不覆盖）
' ============================================================
Sub Move_file(ByVal srcPath As String, ByVal srcFile As String, _
              ByVal outputPath As String, ByVal moveType As String)

    If Right(srcPath, 1) <> "\" And srcPath <> "" Then srcPath = srcPath & "\"
    If Right(outputPath, 1) <> "\" And outputPath <> "" Then outputPath = outputPath & "\"

    If folder_exist(outputPath) = False Then MkDir outputPath

    If file_exist(srcPath & srcFile) And Not workbook_is_open(srcFile) Then
        If file_exist(outputPath & srcFile) Then
            If moveType = "OVERWRITE" Then
                FileCopy srcPath & srcFile, outputPath & srcFile
                Kill srcPath & srcFile
            End If
        Else
            Name srcPath & srcFile As outputPath & srcFile
        End If
    End If

End Sub


' ============================================================
' 复制文件到目标文件夹
' 参数：copyType — "OVERWRITE"（覆盖）或 ""（不覆盖）
' ============================================================
Sub Copy_file(ByVal srcPath As String, ByVal srcFile As String, _
              ByVal outputPath As String, ByVal copyType As String)

    If Right(srcPath, 1) <> "\" And srcPath <> "" Then srcPath = srcPath & "\"
    If Right(outputPath, 1) <> "\" And outputPath <> "" Then outputPath = outputPath & "\"

    If folder_exist(outputPath) = False Then MkDir outputPath

    If file_exist(srcPath & srcFile) Then
        If copyType = "OVERWRITE" Then
            FileCopy srcPath & srcFile, outputPath & srcFile
        ElseIf Not file_exist(outputPath & srcFile) Then
            FileCopy srcPath & srcFile, outputPath & srcFile
        End If
    End If

End Sub


' ============================================================
' 复制整个文件夹（含子文件夹）到目标路径
' ============================================================
Sub Copy_all_files_subfolders(ByVal fromPath As String, ByVal toPath As String)

    Dim fso As Object

    If Right(fromPath, 1) = "\" Then fromPath = Left(fromPath, Len(fromPath) - 1)
    If Right(toPath, 1) = "\" Then toPath = Left(toPath, Len(toPath) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(fromPath) Then
        MsgBox fromPath & " doesn't exist"
        Exit Sub
    End If

    Dim folderName As String
    folderName = Mid(fromPath, InStrRev(fromPath, "\") + 1)
    MkDir toPath & "\" & folderName
    fso.CopyFolder Source:=fromPath, Destination:=toPath & "\" & folderName

End Sub


' ============================================================
' 重命名文件
' ============================================================
Sub Rename_file(ByVal srcPath As String, ByVal srcFile As String, ByVal newFileName As String)

    If Right(srcPath, 1) <> "\" And srcPath <> "" Then srcPath = srcPath & "\"

    If file_exist(srcPath & srcFile) And Not workbook_is_open(srcFile) Then
        Name srcPath & srcFile As srcPath & newFileName
    End If

End Sub


' ============================================================
' 删除文件夹（含所有子文件夹和文件）
' 注意：请确保文件夹内无已打开文件
' ============================================================
Sub Delete_folder(ByVal srcPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right(srcPath, 1) = "\" Then srcPath = Left(srcPath, Len(srcPath) - 1)

    If Not fso.FolderExists(srcPath) Then
        MsgBox srcPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    fso.DeleteFile srcPath & "\*.*", True
    fso.DeleteFolder srcPath & "\*.*", True
    fso.DeleteFolder srcPath
    On Error GoTo 0

End Sub


' ============================================================
' 删除单个文件
' ============================================================
Sub Delete_file(ByVal srcPath As String, ByVal srcFile As String)

    If Right(srcPath, 1) = "\" Then srcPath = Left(srcPath, Len(srcPath) - 1)

    On Error Resume Next
    Kill srcPath & "\" & srcFile
    On Error GoTo 0

End Sub


' ============================================================
' 删除工作簿中的空白工作表
' ============================================================
Sub DelBlankSheet(ByVal wkbk As String)

    Workbooks(wkbk).Activate
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If Application.CountA(ws.UsedRange.Cells) = 0 Then
            ws.Delete
        End If
    Next ws

End Sub


' ============================================================
' 将指定范围内的公式/格式转为静态值（针对特定单元格范围）
' ============================================================
Sub Paste_cell_as_value(ByVal wkbk As String, ByVal wksheet As String, _
                         ByVal pasteRow As String, ByVal pasteCol As String)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Activate
                Call PasteCellAsValueOnSheet(pasteRow, pasteCol)
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Call PasteCellAsValueOnSheet(pasteRow, pasteCol)
    End If

End Sub


' ============================================================
' 私有辅助过程：在当前工作表上执行区域值粘贴
' ============================================================
Private Sub PasteCellAsValueOnSheet(ByVal pasteRow As String, ByVal pasteCol As String)

    Dim rowFrom As Long, rowTo As Long
    Dim colFrom As Long, colTo As Long
    Dim sepPos As Long

    ' 解析行范围
    sepPos = InStr(1, pasteRow, ":")
    rowFrom = CLng(Left(pasteRow, sepPos - 1))
    If UCase(Mid(pasteRow, sepPos + 1)) = "LAST" Then
        rowTo = last_row_check(rowFrom)
    Else
        rowTo = CLng(Mid(pasteRow, sepPos + 1))
    End If

    ' 解析列范围
    sepPos = InStr(1, pasteCol, ":")
    colFrom = Range(Left(pasteCol, sepPos - 1) & "1").Column
    If UCase(Mid(pasteCol, sepPos + 1)) = "LAST" Then
        colTo = last_col_check(rowFrom)
    Else
        colTo = Range(Mid(pasteCol, sepPos + 1) & "1").Column
    End If

    With Range(Cells(rowFrom, colFrom), Cells(rowTo, colTo))
        .Copy
        .PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    Range("A1").Select

End Sub


' ============================================================
' 将整个工作表内容转为静态值
' ============================================================
Sub Paste_sheet_as_value(ByVal wkbk As String, ByVal wksheet As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Activate
                Cells.Copy
                ActiveCell.PasteSpecial xlPasteValues
                Range("A1").Select
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Cells.Copy
        ActiveCell.PasteSpecial xlPasteValues
        Range("A1").Select
    End If

    Application.CutCopyMode = False

End Sub


' ============================================================
' 通过 Outlook 发送电子邮件
' 参数：headerRow — Email 工作表中表头所在行
' ============================================================
Sub Send_Email(ByVal headerRow As Variant)

    Application.DisplayAlerts = False

    Dim i As Integer, j As Integer

    If wbThis Is Nothing Then Set wbThis = ThisWorkbook
    wbThis.Worksheets("Email").Activate

    Dim emailStart As String
    emailStart = "A" & headerRow

    Dim emailCount As Long, emailLastCol As Long
    emailCount = Range(emailStart).End(xlDown).Row - Range(emailStart).Row
    emailLastCol = Range(emailStart).End(xlToRight).Column

    ' 读取邮件列表到数组
    Dim emailList() As Variant
    ReDim emailList(emailCount, emailLastCol)
    For i = 1 To emailCount
        Range(emailStart).Offset(i, 0).Select
        For j = 1 To emailLastCol
            emailList(i, j) = Selection.Value
            Selection.Offset(0, 1).Select
        Next j
    Next i

    ' 逐封发送邮件
    For i = 1 To emailCount
        On Error GoTo ErrHandler

        Dim objOutlook As Object
        Set objOutlook = CreateObject("Outlook.Application")

        Dim objEmail As Object
        Set objEmail = objOutlook.CreateItem(olMailItem)

        With objEmail
            .To = emailList(i, 1)
            .CC = emailList(i, 2)
            .BCC = emailList(i, 3)
            .Subject = emailList(i, 4)
            .Body = emailList(i, 5)
            .Send
        End With

        Set objEmail = Nothing
        Set objOutlook = Nothing

ErrHandler:
    Next i

End Sub


' ============================================================
' 展开工作表中的分组行/列（显示所有层级）
' ============================================================
Sub Expand_group(ByVal wkbk As String, ByVal wksheet As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Activate
                ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
    End If

End Sub


' ============================================================
' 折叠工作表中的分组行/列（收起至第1层）
' ============================================================
Sub Collapse_group(ByVal wkbk As String, ByVal wksheet As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Activate
                ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    End If

End Sub


' ============================================================
' 保护工作表（可带密码）
' ============================================================
Sub Protect_sheet(ByVal wkbk As String, ByVal wksheet As String, ByVal pw As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Protect Password:=pw
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Protect Password:=pw
    End If

End Sub


' ============================================================
' 取消保护工作表（可带密码）
' ============================================================
Sub Unprotect_sheet(ByVal wkbk As String, ByVal wksheet As String, ByVal pw As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name

    If UCase(wksheet) = "ALL" Then
        Dim ws As Worksheet
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                ws.Unprotect Password:=pw
            End If
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Unprotect Password:=pw
    End If

End Sub


' ============================================================
' 保护整个工作簿（可带密码）
' ============================================================
Sub Protect_workbook(ByVal wkbk As String, ByVal pw As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Protect Password:=pw

End Sub


' ============================================================
' 取消保护整个工作簿（可带密码）
' ============================================================
Sub Unprotect_workbook(ByVal wkbk As String, ByVal pw As String)

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Unprotect Password:=pw

End Sub


' ============================================================
' 开启自动计算模式
' ============================================================
Sub Turn_on_auto_cal()
    Application.Calculation = xlAutomatic
End Sub


' ============================================================
' 关闭自动计算模式（切换为手动计算）
' ============================================================
Sub Turn_off_auto_cal()
    Application.Calculation = xlManual
End Sub


' ============================================================
' 删除重复行（按指定列范围）
' 参数：filter_by 格式：起始列字母:结束列字母（如 "A:C"）
' ============================================================
Sub Remove_Duplicate(ByVal wkbk As String, ByVal wksheet As String, _
                      ByVal headerRow As Variant, ByVal filterBy As String)

    Dim i As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Activate

    If wksheet = "" Then wksheet = ActiveSheet.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    Dim lastDataRow As Long
    lastDataRow = last_row_check(headerRow)

    Dim sepPos As Long
    sepPos = InStr(1, filterBy, ":", vbTextCompare)
    Dim colStart As String, colEnd As String
    colStart = Trim(Left(filterBy, sepPos - 1))
    colEnd = Trim(Mid(filterBy, sepPos + 1))

    Dim totalCols As Long
    totalCols = Range(colEnd & headerRow).Column - Range(colStart & headerRow).Column + 1

    Dim colIndexes() As Variant
    ReDim colIndexes(totalCols - 1)
    For i = 0 To totalCols - 1
        colIndexes(i) = i + 1
    Next i

    Range(colStart & headerRow & ":" & colEnd & lastDataRow).RemoveDuplicates _
        Columns:=colIndexes, Header:=xlYes

End Sub


' ============================================================
' 列出工作簿内所有工作表名称到 File 工作表
' ============================================================
Sub List_Worksheet(ByVal srcPath As String, ByVal srcFile As String)

    Dim thisBk As String
    thisBk = ThisWorkbook.Name

    Workbooks(thisBk).Worksheets("File").Activate
    Dim lastDataRow As Long
    lastDataRow = last_row_check(2)
    If lastDataRow = 2 Then lastDataRow = lastDataRow + 1

    Range("A" & lastDataRow).Value = srcPath
    Range("B" & lastDataRow).Value = srcFile

    Dim i As Long
    For i = 2 To lastDataRow - 1
        If Right(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
        If (srcPath & srcFile) <> "" Then
            If Not workbook_is_open(srcFile) Then Call Open_File(srcPath, srcFile, "READ")
        Else
            srcFile = ThisWorkbook.Name
        End If

        Workbooks(thisBk).Worksheets("File").Activate
        Range("B" & i + 1).Select

        Dim ws As Worksheet
        For Each ws In Workbooks(srcFile).Worksheets
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = ws.Name
        Next ws
    Next i

End Sub


' ============================================================
' 对数据透视表的页字段进行筛选
' 参数：filter_by 格式：字段名|值1,值2;字段名2|值...
' ============================================================
Sub Pivot_Filter(ByVal wkbk As String, ByVal wksheet As String, ByVal filterBy As String)

    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim i As Long, j As Long, k As Long

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Set ws = Worksheets(wksheet)

    Dim filterDetails() As String
    filterDetails = Split(filterBy, ";")

    For i = 0 To UBound(filterDetails)
        Dim sepPos As Long
        sepPos = InStr(1, filterDetails(i), "|", vbTextCompare)
        Dim filterField As String
        filterField = Trim(Left(filterDetails(i), sepPos - 1))
        Dim filterValues() As String
        filterValues = Split(Trim(Mid(filterDetails(i), sepPos + 1)), ",")

        For j = 0 To UBound(filterValues)
            If Trim(filterValues(j)) = "(blank)" Then filterValues(j) = ""
        Next j

        For Each PT In ws.PivotTables
            With ws.PivotTables(PT.Name).PivotFields(filterField)
                .ClearAllFilters
                .CurrentPage = "(All)"

                ' 隐藏所有条目，再按需显示
                For j = 1 To .PivotItems.Count - 1
                    .PivotItems(j).Visible = False
                Next j

                For j = 1 To .PivotItems.Count - 1
                    For k = 0 To UBound(filterValues)
                        If UCase(.PivotItems(j).Name) = UCase(filterValues(k)) Then
                            .PivotItems(j).Visible = True
                        End If
                    Next k
                Next j

                .PivotItems("(blank)").Visible = False
                .EnableMultiplePageItems = True
            End With
            PT.RefreshTable
        Next PT
    Next i

End Sub


' ============================================================
' 刷新数据库连接和数据透视表，并记录刷新时间
' ============================================================
Sub Data_RefreshAll(ByVal wkbk As String, ByVal wksheet As String)

    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim db As WorkbookConnection
    Dim i As Long

    If wkbk = "" Then wkbk = ThisWorkbook.Name

    If UCase(wksheet) = "ALL" Then
        For Each ws In Workbooks(wkbk).Worksheets
            If Not IsCtrlSheet(ws.Name) Then
                For Each PT In ws.PivotTables
                    PT.SourceData = check_pivot_range(PT.SourceData)
                    PT.RefreshTable
                Next PT
            End If
        Next ws
    Else
        For Each PT In Workbooks(wkbk).Worksheets(wksheet).PivotTables
            PT.RefreshTable
        Next PT
    End If

    ' 记录查询连接的最后刷新时间
    i = wbThis.Worksheets("Query_LastRefresh").Range("A1").End(xlDown).Row
    If i = 1048576 Then i = 1

    Dim lastRefresh As String
    For Each db In Workbooks(wkbk).Connections
        On Error Resume Next
        Workbooks(wkbk).Connections(db.Name).Refresh
        If Err.Number = 1004 Then
            lastRefresh = "Can't access to data base"
        Else
            lastRefresh = Now
        End If
        On Error GoTo 0
        wbThis.Worksheets("Query_LastRefresh").Range("A" & i + 1) = db.Name
        wbThis.Worksheets("Query_LastRefresh").Range("B" & i + 1) = lastRefresh
        i = i + 1
    Next db

End Sub


' ============================================================
' 复制源工作表的全部数据（从A1开始）到目标工作表
' ============================================================
Sub Copy_sheetdata(ByVal wkbk As String, ByVal wksheet As String, _
                   ByVal srcPath As String, ByVal srcFile As String, _
                   ByVal srcSheet As String, ByVal srcHeaderRow As Variant, _
                   ByVal srcDataCol As Variant)

    Application.DisplayAlerts = False

    If srcFile = "N/A" Then Exit Sub

    If (srcPath & srcFile) <> "" Then
        If Not workbook_is_open(srcFile) Then Call Open_File(srcPath, srcFile, "READ")
    Else
        srcFile = ThisWorkbook.Name
    End If

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    If srcSheet = "" Then srcSheet = Workbooks(srcFile).ActiveSheet.Name
    Workbooks(srcFile).Worksheets(srcSheet).Activate

    Dim srcLastCol As Long, srcLastRow As Long
    srcLastCol = last_col_check(srcHeaderRow)
    srcLastRow = last_row_check(srcHeaderRow)

    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)

    ' 复制源工作表全部数据区域并粘贴为值
    Workbooks(srcFile).Worksheets(srcSheet).Range( _
        Workbooks(srcFile).Worksheets(srcSheet).Cells(1, 1), _
        Workbooks(srcFile).Worksheets(srcSheet).Cells(srcLastRow, srcLastCol)).Copy
    Workbooks(wkbk).Worksheets(wksheet).Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats

    Range("A1").Select
    Application.CutCopyMode = False

    If (srcPath & srcFile) <> "" And srcFile <> ThisWorkbook.Name Then
        Call Close_File(srcPath, srcFile, "UNCHANGE")
    End If

End Sub


' ============================================================
' 复制工作表到目标工作簿
' ============================================================
Sub Copy_Sheet(ByVal srcPath As String, ByVal srcFile As String, _
               ByVal srcSheet As String, ByVal outputPath As String, ByVal outputFile As String)

    Application.DisplayAlerts = False

    If srcFile = "" Then srcFile = ActiveWorkbook.Name
    If srcPath = "" Then srcPath = ActiveWorkbook.Path
    If Right(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"

    If Not workbook_is_open(srcFile) Then Call Open_File(srcPath, srcFile, "READ")

    ' 若目标工作簿不存在则新建
    Dim needDeleteBlank As Boolean
    needDeleteBlank = False
    If file_exist(outputPath & outputFile) Then
        Call Open_File(outputPath, outputFile, "EDIT")
    Else
        Workbooks.Add
        ActiveWorkbook.Sheets(1).Name = "~"
        ActiveWorkbook.SaveAs outputPath & outputFile
        needDeleteBlank = True
    End If
    Dim insertBefore As String
    insertBefore = ActiveSheet.Name

    ' 支持"|"分隔的多工作表复制
    Dim sheetList() As String
    sheetList = Split(srcSheet, "|")
    Workbooks(srcFile).Activate

    Dim i As Integer
    For i = 0 To UBound(sheetList)
        Workbooks(srcFile).Sheets(sheetList(i)).Copy Before:=Workbooks(outputFile).Sheets(insertBefore)
    Next i

    ' 删除新建工作簿时创建的占位空白工作表
    If needDeleteBlank Then
        Workbooks(outputFile).Worksheets("~").Delete
    End If

    Call Close_File(outputPath, outputFile, "SAVE")

End Sub


' ============================================================
' 删除工作表
' 参数：wk_bk — 工作簿名称；delBy — 工作表名（"|"分隔多个）
' 【Bug修复】原代码误用未赋值的 wkbk 变量，已改为参数 wk_bk
' ============================================================
Sub Delete_sheet(ByVal wk_bk As String, ByVal delBy As String)

    If wk_bk = "" Then wk_bk = ThisWorkbook.Name
    Workbooks(wk_bk).Activate

    Dim sheetList() As String
    sheetList = Split(delBy, "|")

    Dim i As Integer
    For i = 0 To UBound(sheetList)
        If sheet_exist(wk_bk, Trim(sheetList(i))) Then
            Workbooks(wk_bk).Worksheets(Trim(sheetList(i))).Delete
        End If
    Next i

End Sub


' ============================================================
' 从 Cut file 工作表读取配置，对目标文件执行裁剪操作
' （按条件筛选删除不符合要求的行，并执行列删除、工作表删除、透视表刷新）
' ============================================================
Sub Cut_File()

    Application.DisplayAlerts = False

    Dim i As Integer, j As Integer, k As Integer
    Const CRITERIA_FIRST_COL As Integer = 9   ' 筛选条件从第9列开始

    If wbThis Is Nothing Then Set wbThis = ThisWorkbook
    wbThis.Worksheets("Cut file").Activate

    Dim cutStartCell As String
    cutStartCell = Range("B1").Value

    Dim cutFileCount As Long, cutFileLastCol As Long
    cutFileCount = Range(cutStartCell).End(xlDown).Row - Range(cutStartCell).Row
    cutFileLastCol = Range(cutStartCell).End(xlToRight).Column

    ' 读取裁剪配置到数组
    Dim cutFileData() As Variant
    ReDim cutFileData(cutFileCount, cutFileLastCol)
    For i = 1 To cutFileCount
        Range(cutStartCell).Offset(i, 0).Select
        For j = 1 To cutFileLastCol
            cutFileData(i, j) = Selection.Value
            Selection.Offset(0, 1).Select
        Next j
    Next i

    ' 读取筛选条件表头
    Dim criteria() As Variant
    ReDim criteria(cutFileLastCol - CRITERIA_FIRST_COL)
    Range(cutStartCell).Offset(0, CRITERIA_FIRST_COL).Select
    For i = 1 To cutFileLastCol - CRITERIA_FIRST_COL
        criteria(i) = Selection.Value
        Selection.Offset(0, 1).Select
    Next i

    ' 逐文件处理
    For i = 1 To cutFileCount
        If Right(cutFileData(i, 1), 1) <> "\" Then cutFileData(i, 1) = cutFileData(i, 1) & "\"
        If Right(cutFileData(i, 5), 1) <> "\" Then cutFileData(i, 5) = cutFileData(i, 5) & "\"

        ' 若目标文件不存在则从源复制
        If Not file_exist(cutFileData(i, 5) & cutFileData(i, 6)) Then
            FileCopy cutFileData(i, 1) & cutFileData(i, 2), cutFileData(i, 5) & cutFileData(i, 6)
        End If

        Call Open_File(cutFileData(i, 5), cutFileData(i, 6), "EDIT")
        Dim currentWs As String
        currentWs = ActiveSheet.Name
        Workbooks(cutFileData(i, 6)).Activate

        If cutFileData(i, 3) = "" Then cutFileData(i, 3) = ActiveSheet.Name
        Worksheets(cutFileData(i, 3)).Activate

        ' 记录并临时清除筛选状态
        Dim filterFlag As Integer
        filterFlag = 0
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
            filterFlag = 1
        End If

        ' 按条件筛选并删除不符合的行
        For j = 1 To UBound(criteria)
            If cutFileData(i, j + CRITERIA_FIRST_COL) <> "" Then
                ' 在表头下方插入空行（避免合并单元格问题）
                Rows(cutFileData(i, 4) + 1 & ":" & cutFileData(i, 4) + 1).Insert _
                    Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Rows(cutFileData(i, 4) + 1 & ":" & cutFileData(i, 4) + 1).AutoFilter

                Range("A" & cutFileData(i, 4)).Select
                For k = 1 To last_col_check(cutFileData(i, 4))
                    If Selection.Value = criteria(j) Then
                        Dim multiSelect As Variant
                        If InStr(cutFileData(i, j + CRITERIA_FIRST_COL), ";") > 0 Then
                            multiSelect = Split(cutFileData(i, j + CRITERIA_FIRST_COL), ";")
                            Dim m As Integer
                            For m = 0 To UBound(multiSelect)
                                multiSelect(m) = Trim(multiSelect(m))
                            Next m
                            ' 多值筛选（最多支持2个排除值）
                            Selection.AutoFilter Field:=Selection.Column, _
                                Criteria1:="<>" & multiSelect(0), Operator:=xlAnd, _
                                Criteria2:="<>" & multiSelect(1)
                        Else
                            ' 单值筛选
                            Selection.AutoFilter Field:=Selection.Column, _
                                Criteria1:="<>" & cutFileData(i, j + CRITERIA_FIRST_COL)
                        End If

                        ' 删除插入的临时行和筛选出来的不符合行
                        Rows(cutFileData(i, 4) + 1 & ":" & cutFileData(i, 4) + 1).Select
                        Range(Selection, Selection.End(xlDown).End(xlDown)).Delete Shift:=xlUp

                        Exit For
                    End If
                    Selection.Offset(0, 1).Select
                Next k
            End If
        Next j

        Range("A1").Select

        ' 还原原有筛选状态
        If filterFlag = 1 Then
            Rows(cutFileData(i, 4) & ":" & cutFileData(i, 4)).AutoFilter
            Range("A1").Select
        End If

        ' 删除指定列/行
        If cutFileData(i, 7) <> "" Then
            Call Delete_Col_Row(cutFileData(i, 6), cutFileData(i, 3), cutFileData(i, 7))
            Range("A1").Select
        End If

        ' 删除指定工作表
        If cutFileData(i, 8) <> "" Then
            Call Delete_sheet(cutFileData(i, 6), cutFileData(i, 8))
        End If

        ' 刷新数据透视表
        If cutFileData(i, 9) <> "" Then
            Call Refresh_Pivot(cutFileData(i, 6), UCase(cutFileData(i, 9)))
        End If

        If sheet_exist(cutFileData(i, 6), currentWs) Then Range("A1").Select

        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Next i

End Sub


' ============================================================
' 将日期格式从 dd.mm.yyyy 转换为 dd/mm/yyyy
' ============================================================
Sub Change_dot_to_slash(ByVal wkbk As String, ByVal wksheet As String, _
                         ByVal headerRow As Variant, ByVal startRow As Variant, _
                         ByVal dotCol As String)

    Dim i As Integer, j As Integer

    If wkbk = "" Then wkbk = ThisWorkbook.Name
    Workbooks(wkbk).Worksheets(wksheet).Activate

    If startRow = "" Then startRow = headerRow + 1

    Dim lastDataRow As Long
    lastDataRow = last_row_check(headerRow)

    Dim colList() As String
    colList = Split(dotCol, "|")

    For i = 0 To UBound(colList)
        For j = startRow To lastDataRow
            Dim cell As Range
            Set cell = Range(colList(i) & j)
            cell.Value = DateValue(Left(cell.Value, 2) & "/" & Mid(cell.Value, 4, 2) & "/" & Right(cell.Value, 4))
        Next j
        Range(colList(i) & startRow & ":" & colList(i) & lastDataRow).NumberFormat = "dd/mm/yyyy"
    Next i

End Sub


' ============================================================
' 【性能优化】获取当前工作表最后使用列号
' 原实现遍历所有列（性能极差），改用 UsedRange 快速定位
' ============================================================
Function last_col_check(ByVal headerRow As Variant) As Long

    last_col_check = ActiveSheet.UsedRange.Columns.Count + ActiveSheet.UsedRange.Column - 1

End Function


' ============================================================
' 【性能优化】获取当前工作表最后使用行号
' 原实现遍历所有行（性能极差），改用 UsedRange 快速定位
' ============================================================
Function last_row_check(ByVal headerRow As Variant) As Long

    last_row_check = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Row - 1

End Function


' ============================================================
' 获取最后使用区域的右下角单元格地址（如 "H20"）
' ============================================================
Function last_cell_check(ByVal headerRow As Variant) As String

    Dim lastColCell As String
    lastColCell = Cells(headerRow, Columns.Count).End(xlToLeft).Address
    Dim lastDataRow As Long
    lastDataRow = last_row_check(headerRow)
    ' 提取列字母部分（去掉 $ 符号），拼接行号
    last_cell_check = Left(lastColCell, InStr(2, lastColCell, "$", vbTextCompare) - 1) & lastDataRow

End Function


' ============================================================
' 判断当前选中单元格是否属于数据透视表
' ============================================================
Function cell_in_pivot() As Boolean

    Dim PT As PivotTable
    On Error Resume Next
    Set PT = ActiveCell.PivotTable
    On Error GoTo 0

    cell_in_pivot = Not (PT Is Nothing)

End Function


' ============================================================
' 判断文件夹是否存在
' ============================================================
Function folder_exist(ByVal fullPath As String) As Boolean
    folder_exist = (Dir(fullPath, vbDirectory) <> vbNullString)
End Function


' ============================================================
' 判断文件是否存在
' ============================================================
Function file_exist(ByVal fullPath As String) As Boolean
    file_exist = (Dir(fullPath) <> "")
End Function


' ============================================================
' 判断指定工作表是否存在于工作簿中
' ============================================================
Function sheet_exist(ByVal wb As String, ByVal ws As String) As Boolean
    On Error Resume Next
    sheet_exist = (Workbooks(wb).Sheets(ws).Index > 0)
    On Error GoTo 0
End Function


' ============================================================
' 判断工作簿是否已在 Excel 中打开
' ============================================================
Function workbook_is_open(ByVal wbName As String) As Boolean

    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbName)
    workbook_is_open = (Err.Number = 0)
    On Error GoTo 0

End Function


' ============================================================
' 检查数据透视表源数据是否只有表头行，若是则额外加一行以避免透视刷新报错
' ============================================================
Function check_pivot_range(ByVal pivotRange As String) As String

    Dim sheetPart As String, cell1 As String, cell2 As String
    Dim row1 As String, row2 As String

    sheetPart = Left(pivotRange, InStr(pivotRange, "!"))
    cell1 = Mid(pivotRange, InStr(pivotRange, "!") + 2, InStr(pivotRange, ":") - InStr(pivotRange, "!") - 1)
    cell2 = Right(pivotRange, Len(pivotRange) - InStr(pivotRange, ":") - 1)

    row1 = Left(cell1, InStr(cell1, "C") - 1)
    row2 = Left(cell2, InStr(cell2, "C") - 1)

    If row1 = row2 Then
        ' 仅有表头行时，将结束行+1
        row2 = CStr(CLng(row2) + 1)
        check_pivot_range = sheetPart & "R" & cell1 & "R" & row2 & "C" & Mid(cell2, InStr(cell2, "C") + 1)
    Else
        check_pivot_range = pivotRange
    End If

End Function


' ============================================================
' 私有辅助函数：判断工作表名称是否在控制表白名单中
' 用于所有 "ALL" 模式操作跳过控制工作表
' ============================================================
Private Function IsCtrlSheet(ByVal sheetName As String) As Boolean

    Dim i As Integer
    For i = 0 To UBound(ctrlSheetNames)
        If sheetName = ctrlSheetNames(i) Then
            IsCtrlSheet = True
            Exit Function
        End If
    Next i
    IsCtrlSheet = False

End Function
