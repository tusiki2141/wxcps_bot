Dim wbThis As Workbook
Dim wsThis As Worksheet
Dim Action_seq As Variant
Dim ctrl_no, ctrl_lastrow As Integer
Dim ctrl_err As String
Dim Cnt_open, Cnt_close As Integer
Dim ctrl_ws As Variant


Sub Combination_main()
'Main Function
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Beg = Now
    
    'Define Control worksheet
    ctrl_ws = Array("Control", "Src", "File", "Cut file", "Email")
    
    'Define workbook & control worksheet
    ThisWorkbook.Activate
    Set wbThis = ThisWorkbook
    Worksheets("Control").Activate
    Set wsThis = wbThis.ActiveSheet
    input_session = Range("B1").Value
    
    'Split session start cell for multiple input
    exe_sec = Split(input_session, "|")
    
    'Run each session
    For i = 0 To UBound(exe_sec)
        Run_action (exe_sec(i))
    Next i
    
    'Show Macro workbook and Control Worksheet
    wbThis.Activate
    wsThis.Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Range("B1").Select
    MsgBox ("Completed. " & vbNewLine & Beg & vbNewLine & Now)
   
End Sub

Sub Run_action(exe_sec)
'Run actions/steps, based on user input (session start cell & control table)

    'Get control table data
    Get_control (exe_sec)
    
    'Validate Control table input
    Validate_Control (exe_sec)
    
    action_name = Range(exe_sec).Offset(-2, 0).Value
    row_no = Range(exe_sec).Offset(1, 0).Row
    
    'Check action type and call corresponding function
    For i = 1 To ctrl_no
    
        Application.ScreenUpdating = False
        Application.StatusBar = "Processing..." & action_name & " Step " & i & " - Row_" & row_no & "_" & Action_seq(i, 1) & ". Total " & FormatPercent(i / ctrl_no) & " Completed."
        row_no = row_no + 1
        If (UCase(Trim(Action_seq(i, 1))) = "CLEAR_DATA") Then
            'parameter(wkbk, wksheet, header_row, start_row)
            Call Clear_Data(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 7))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "COPY_FORMULA") Then
            'parameter(wkbk, wksheet, ind_row, header_row, formula_row, start_row)
            Call Copy_Formula(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 4), Action_seq(i, 6), Action_seq(i, 5), Action_seq(i, 7))
            
        ElseIf InStr(1, UCase(Trim(Action_seq(i, 1))), "OPEN_FILE", vbTectCompare) > 0 Then
            
            If InStr(1, UCase(Trim(Action_seq(i, 1))), "EDIT", vbTextCompare) > 0 Then
                'parameter(src_path, src_bk, open_type)
                Call Open_File(Action_seq(i, 12), Action_seq(i, 13), "EDIT")
                
            Else
                'parameter(src_path, src_bk, open_type)
                Call Open_File(Action_seq(i, 12), Action_seq(i, 13), "READ")
                
            End If
            
        ElseIf InStr(1, UCase(Trim(Action_seq(i, 1))), "CLOSE_FILE", vbTectCompare) > 0 Then
            
            If InStr(1, UCase(Trim(Action_seq(i, 1))), "SAVE AS", vbTextCompare) > 0 Then
                'parameter(output_path, output_bk, close_type)
                Call Close_File(Action_seq(i, 18), Action_seq(i, 19), "SAVE AS")
                
            ElseIf InStr(1, UCase(Trim(Action_seq(i, 1))), "SAVE", vbTextCompare) > 0 Then
                'parameter(output_path, output_bk, close_type)
                Call Close_File(Action_seq(i, 18), Action_seq(i, 19), "SAVE")
                
            Else
                'parameter(output_path, output_bk, close_type)
                Call Close_File(Action_seq(i, 18), Action_seq(i, 19), "UNCHANGE")
                
            End If
                        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "SAVE_FILE") Then
            'parameter(output_path, output_bk)
            Call Save_File(Action_seq(i, 18), Action_seq(i, 19))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "APPEND_ALL") Then
            'parameter(wkbk, wksheet, ind_row, header_row, add_srctext,src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
            Call Append_All(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 4), Action_seq(i, 6), Action_seq(i, 10), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 16), Action_seq(i, 17))
                   
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "APPEND_ALL_NOTCLOSEFILE") Then
            'parameter(wkbk, wksheet, ind_row, header_row, add_srctext,src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
            Call Append_All_notCloseFile(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 4), Action_seq(i, 6), Action_seq(i, 10), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 16), Action_seq(i, 17))
                   
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "APPEND_BY_COL_NAME") Then
            'parameter(wkbk, wksheet, header_row,add_srctext, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
            Call Append_by_Col_Name(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 10), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 16), Action_seq(i, 17))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "APPEND_BY_COL_NAME_NOTCLOSEFILE") Then
            'parameter(wkbk, wksheet, header_row, add_srctext,src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
            Call Append_by_Col_Name_notCloseFile(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 10), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 16), Action_seq(i, 17))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "APPEND_IN_SAME_LINE") Then
            'parameter(wkbk, wksheet,indicator, header_row, start_row, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
            Call Append_in_Same_Line(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 4), Action_seq(i, 6), Action_seq(i, 7), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 16), Action_seq(i, 17))
                
                
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "REFRESH_PIVOT") Then
            'parameter(wkbk, wksheet)
            Call Refresh_Pivot(Action_seq(i, 2), Action_seq(i, 3))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "FILTER") Then
            'parameter(wkbk, wksheet, header_row, filter_by)
            Call Filter(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 9))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "FILTER_PERIOD") Then
            'parameter(wkbk, wksheet, header_row, filter_by)
            Call Filter_Period(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 9))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "SORTING") Then
            'parameter(wkbk, wksheet, header_row, sort_by)
            Call Sorting(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 8))
                
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "UNFILTER") Then
            'parameter(wkbk, wksheet)
            Call Unfilter(Action_seq(i, 2), Action_seq(i, 3))
           
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "ADD_TEXT") Then
            'parameter(wkbk, wksheet, header_row, start_row, text_to_add)
            Call add_text(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 7), Action_seq(i, 10))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "CHANGE_TEXT") Then
            'parameter(wkbk, wksheet, header_row, start_row, text_to_add)
            Call Change_text(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 7), Action_seq(i, 10))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "DELETE_COL_ROW") Then
            'parameter(wkbk, wksheet, del_by)
            Call Delete_Col_Row(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 11))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "LIST_FILE") Then
            'parameter(src_path, src_type, wksheet)
            Call List_file(Action_seq(i, 12), Action_seq(i, 13), "FILE")
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "LIST_FILES") Then
            'parameter(src_path, src_type, wksheet)
            Call List_files(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 4), Action_seq(i, 15))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "LIST_SUBFOLDER") Then
            'parameter(src_path, wksheet)
            Call List_subfolder(Action_seq(i, 12), "File")
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "LIST_ALL_FILES_SUBFOLDERS") Then
            'parameter(src_path, wksheet)
            Call List_all_files_subfolders(Action_seq(i, 12), "File")
        
        ElseIf InStr(1, UCase(Trim(Action_seq(i, 1))), "MOVE_FILE", vbTectCompare) > 0 Then
        
            If InStr(1, UCase(Trim(Action_seq(i, 1))), "OVERWRITE", vbTextCompare) > 0 Then
                'parameter(src_path, src_bk, output_path, move_type)
                Call Move_file(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 18), "OVERWRITE")
                
            Else
                'parameter(src_path, src_bk, output_path, move_type)
                Call Move_file(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 18), "")
                
            End If
        
        ElseIf InStr(1, UCase(Trim(Action_seq(i, 1))), "COPY_FILE", vbTectCompare) > 0 Then

            If InStr(1, UCase(Trim(Action_seq(i, 1))), "OVERWRITE", vbTextCompare) > 0 Then
                'parameter(src_path, src_bk, output_path, copy_type)
                Call Copy_file(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 18), "OVERWRITE")
                
            Else
                'parameter(src_path, src_bk, output_path, copy_type)
                Call Copy_file(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 18), "")
                
            End If
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "COPY_ALL_FILES_SUBFOLDERS") Then
            'parameter(fromPath, toPath)
            Call Copy_all_files_subfolders(Action_seq(i, 12), Action_seq(i, 18))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "RENAME_FILE") Then
            'parameter(src_path, src_bk, output_bk)
            Call Rename_file(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 19))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "COPY_SHEET") Then
            'parameter(src_path, src_bk, src_sheet, output_path, output_bk)
            Call Copy_Sheet(Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 18), Action_seq(i, 19))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "DELETE_SHEET") Then
            'parameter(wk_bk, del_by)
            Call Delete_sheet(Action_seq(i, 2), Action_seq(i, 11))
             
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "CUT_FILE") Then
            'parameter()
            Call Cut_File
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "PASTE_SHEET_AS_VALUE") Then
            'parameter(wkbk, wksheet)
            Call Paste_sheet_as_value(Action_seq(i, 2), Action_seq(i, 3))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "PASTE_CELL_AS_VALUE") Then
            'parameter(wkbk, wksheet, paste_row, paste_col)
            Call Paste_cell_as_value(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 16), Action_seq(i, 17))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "SEND_EMAIL") Then
            'parameter(header_row)
            Call Send_Email(Action_seq(i, 6))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "DELETE_FOLDER") Then
            'parameter(src_path)
            Call Delete_folder(Action_seq(i, 12))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "DELETE_FILE") Then
            'parameter(src_path, src_bk)
            Call Delete_file(Action_seq(i, 12), Action_seq(i, 13))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "EXPAND_GROUP") Then
            'parameter(wkbk, wksheet)
            Call Expand_group(Action_seq(i, 2), Action_seq(i, 3))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "COLLAPSE_GROUP") Then
            'parameter(wkbk, wksheet)
            Call Collapse_group(Action_seq(i, 2), Action_seq(i, 3))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "PROTECT_SHEET") Then
            'Expand_group(wkbk, wksheet, pw)
            Call Protect_sheet(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 10))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "UNPROTECT_SHEET") Then
            'parameter(wkbk, wksheet, pw)
            Call Unprotect_sheet(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 10))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "PROTECT_WORKBOOK") Then
            'parameter(wkbk, pw)
            Call Protect_workbook(Action_seq(i, 2), Action_seq(i, 10))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "UNPROTECT_WORKBOOK") Then
            'parameter(wkbk, pw)
            Call Unprotect_workbook(Action_seq(i, 2), Action_seq(i, 10))
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "TURN_ON_AUTO_CAL") Then
            'parameter()
            Call Turn_on_auto_cal
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "TURN_OFF_AUTO_CAL") Then
            'parameter()
            Call Turn_off_auto_cal
        
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "REMOVE_DUPLICATE") Then
            'parameter(wkbk, pw)
            Call Remove_Duplicate(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 6), Action_seq(i, 9))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "LIST_WORKSHEET") Then
            'parameter(wkbk, pw)
            Call List_Worksheet(Action_seq(i, 12), Action_seq(i, 13))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "PIVOT_FILTER") Then
            'parameter(wkbk, wksheet, header_row, filter_by)
            Call Pivot_Filter(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 9))
            
        ElseIf (UCase(Trim(Action_seq(i, 1))) = "DATA_REFRESHALL") Then
            'parameter(wkbk, wksheet,ind_row)
            Call Data_RefreshAll(Action_seq(i, 2), Action_seq(i, 3))
            
         ElseIf (UCase(Trim(Action_seq(i, 1))) = "COPY_SHEETDATA") Then
            'parameter(wkbk, wksheet,ind_row)
            Call Copy_sheetdata(Action_seq(i, 2), Action_seq(i, 3), Action_seq(i, 12), Action_seq(i, 13), Action_seq(i, 14), Action_seq(i, 15), Action_seq(i, 17))
            
        End If
        Application.StatusBar = "Processing...Step " & i & " - " & Action_seq(i, 1) & ". Total " & FormatPercent(i / ctrl_no) & " Completed."
        'DoEvents
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


Sub Get_control(ctrl_start)
'Get user input in control table

    'Get number of commands, last used columns & rows from Control Table
    wsThis.Activate
    ctrl_no = Range(ctrl_start).End(xlDown).Row - Range(ctrl_start).Row
    ctrl_lastrow = Range(ctrl_start).End(xlDown).Row
    ctrl_lastcol = Range(ctrl_start).End(xlToRight).Column
    
    'Loop Control Table data into an array
    ReDim Action_seq(ctrl_no, ctrl_lastcol)
    Dim i As Integer
    
    For i = 1 To ctrl_no
        Range(ctrl_start).Offset(i, 0).Select
        For j = 1 To ctrl_lastcol
            Action_seq(i, j) = Selection.Value
            Selection.Offset(0, 1).Select
        Next j
    Next i
    
End Sub


Sub Validate_Control(ctrl_start)
'Check if all the mandatory fields have been filled


End Sub

Sub Clear_Data(wkbk, wksheet, header_row, start_row)
'Clean up all the rows after the header row
'For worksheet with pivot table, it will clear those columns does not belongs to a pivot
    
    'Use current workbook if user input is blank
    Dim i As Integer
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    'Check start row and last used row
    If start_row = "" Then start_row = header_row + 1
    last_row = last_row_check(header_row)
    If last_row = header_row Then last_row = last_row + 1
    If last_row < start_row Then last_row = start_row
    
    'Check if any pivot in worksheet
    If Worksheets(wksheet).PivotTables.Count > 0 Then
        last_col = last_col_check(header_row)
        For i = 1 To last_col
            Cells(header_row, i).Select
            'Check if the columns belongs to a pivot table
            If cell_in_pivot() = False Then
                Range(Cells(start_row, i), Cells(last_row, i)).Clear
            End If
        Next i
    Else
        Range(Rows(start_row), Rows(last_row)).EntireRow.Delete
    End If
    
End Sub


Sub Copy_Formula(wkbk, wksheet, ind_row, header_row, formula_row, start_row)
'To copy formula from formula rows down to all records for column indicated as "Formula"

    'Use current workbook if user input is blank
    
    Dim i As Integer
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
        
    'Find the last used header columns & rows
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
            
    'Check start row
    If start_row = "" Then
        If header_row = last_row Then
            Exit Sub
        Else
            For i = header_row + 1 To last_row
                If Not Rows(i).Rows.Hidden Then
                    start_row = i
                    Exit For
                End If
            Next i
        End If
    End If
        
    'Find the indicator "Formula" in worksheet and copy the formula down to rows
    Range("A" & ind_row).Select
    For i = 1 To last_col
        If Selection.MergeCells = False Then
            If UCase(Selection.Value) = "FORMULA" Then
                Cells(formula_row, Selection.Column).Copy
                Range(Cells(start_row, Selection.Column), Cells(last_row, Selection.Column)).PasteSpecial xlPasteFormulasAndNumberFormats
                Call Unfilter(wkbk, wksheet)
                Range(Cells(start_row, Selection.Column), Cells(last_row, Selection.Column)).Select
                Selection.Copy
                ActiveCell.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
            End If
        End If
        Cells(ind_row, i + 1).Select
    Next i
    
End Sub


Sub Open_File(src_path, src_bk, open_type)
'Open workbook as read / Edit

    'Check source file path format
    If Right(src_path, 1) <> "\" Then src_path = src_path & "\"
    
    'Check open file as read or edit
    If open_type = "READ" Then
        Workbooks.Open Filename:=src_path & src_bk, ReadOnly:=True, UpdateLinks:=False
        
    ElseIf open_type = "EDIT" Then
        Workbooks.Open Filename:=src_path & src_bk, UpdateLinks:=False
        
    End If
    
End Sub


Sub Close_File(output_path, output_bk, close_type)
'Close workbook
    
    'Use current path if not specify
    If output_path = "" Then
        output_path = ThisWorkbook.Path & "\"
    End If
    
    If close_type = "SAVE AS" Then
        'Check file path format
        If Right(output_path, 1) <> "\" Then output_path = output_path & "\"
        ActiveWorkbook.SaveAs Filename:=output_path & output_bk
    End If
    
    'Check if file is opened already
    If workbook_is_open(output_bk) Then
    
        If close_type = "SAVE" Then
            Workbooks(output_bk).Close savechanges:=True
                   
        Else
            Workbooks(output_bk).Close savechanges:=False
            
        End If
    End If
       
End Sub

Sub Save_File(output_path, output_bk)
'Close workbook
    
    'Use current path if not specify
    If output_path = "" Then
        output_path = ThisWorkbook.Path & "\"
        
        Else
        If Right(output_path, 1) <> "\" Then output_path = output_path & "\"
        
    End If
    
    'Use current workbook if user input is blank
    If output_bk = "" Then
        output_bk = ThisWorkbook.Name
    End If
    
    'Check if file is opened already
    If workbook_is_open(output_bk) Then
        Workbooks(output_bk).Activate
        ActiveWorkbook.Save
    Else
        Workbooks(ThisWorkbook.Name).SaveAs (output_path & output_bk)
    End If
End Sub


Sub Append_All(wkbk, wksheet, ind_row, header_row, add_srctext, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
'Using "Source_Start" as an indicator to append all source data to working worksheet
'Note: This function does not check the column name.
'      Please make sure the columns are same as the source data

    'Skip this step if user input N/A in source workbook
    Dim i As Integer
    
    If src_bk = "N/A" Then
       Exit Sub
    End If

    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    
    'Find the last used header columns & rows
    Call Unfilter(wkbk, wksheet)
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Find the indicator "Source_Start" in worksheet
    Range("A" & ind_row).Select
    For i = 1 To last_col
        If UCase(Selection.Value) = "SOURCE_START" Then
            Cells(last_row + 1, Selection.Column).Select
            
            'If file path & file name are provided, open file. If not, use current workbook
            If (src_path & src_bk) <> "" Then
                'If file has already open, it will not be re-opened
                If workbook_is_open(src_bk) = False Then
                    Call Open_File(src_path, src_bk, "READ")
                End If
                         
            Else
                src_bk = ThisWorkbook.Name
        
            End If
            
            'If no worksheet name, set current worksheet as source worksheet
            If src_sheet = "" Then
                src_sheet = Workbooks(src_bk).ActiveSheet.Name
            End If
            Workbooks(src_bk).Worksheets(src_sheet).Activate
            
            'Select data to be copied. This can select data with blank cells but separated by whole blank row
            If src_data_col = "" Then src_data_col = "A"
            src_last_cell = last_cell_check(src_header_row)
            src_last_row = last_row_check(src_header_row)

            If src_last_row >= src_data_row Then
                Range(Cells(src_data_row, src_data_col), src_last_cell).Select
                
                'Copy and paste as value
                Selection.Copy
                Workbooks(wkbk).Worksheets(wksheet).Activate
                ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
                Application.CutCopyMode = False
                
                'Give source remark,source from which path, files, worksheet base on text in Control Column J ->"Add fixed text/pw"
    
                If add_srctext <> "" And add_srctext <> "NA" Then
                    add_srctext_details = Split(add_srctext, ";")
                    For k = 0 To UBound(add_srctext_details)
                        add_to_col = Range(Trim(Left(add_srctext_details(k), InStr(1, add_srctext_details(k), "|", vbTextCompare) - 1)) & "1").Column
                        add_value = Split(Trim(Right(add_srctext_details(k), Len(add_srctext_details(k)) - InStr(1, add_srctext_details(k), "|", vbTextCompare))), ",")
                    
                        Range(Cells(last_row + 1, add_to_col), Cells(last_row + src_last_row - src_data_row + 1, add_to_col)).Value = add_value
                    Next k
                End If
                
            End If
            
            'If file path & file name are provided, close file
            If (src_path & src_bk) <> "" And src_bk <> ThisWorkbook.Name Then
                Call Close_File(src_path, src_bk, "UNCHANGE")
            End If
            
            Exit For
        End If
        
        Selection.Offset(0, 1).Select
        
    Next i
    Range("A" & ind_row).Select
End Sub


Sub Append_by_Col_Name(wkbk, wksheet, header_row, add_srctext, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
'To check and copy columns with same name to designated worksheet
    
    'Skip this step if user input N/A in source workbook
    Dim i, j As Integer
    If src_bk = "N/A" Then
       Exit Sub
    End If
    
    
    'If file path & file name are provided, open file. If not, use current workbook
    If (src_path & src_bk) <> "" Then
        'If file has already open, it will not be re-opened
        If workbook_is_open(src_bk) = False Then
            Call Open_File(src_path, src_bk, "READ")
        End If
    Else
        src_bk = ThisWorkbook.Name

    End If
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    'If no worksheet name, set current worksheet as source worksheet
    If src_sheet = "" Then
        src_sheet = Workbooks(src_bk).ActiveSheet.Name
    End If
    Workbooks(src_bk).Worksheets(src_sheet).Activate
    
    'Find the last used header rows & columns of the source worksheet
    If src_data_col = "" Then src_data_col = "A"
    src_start_col = Range(src_data_col & src_header_row).Column
    src_last_col = last_col_check(src_header_row)
    src_last_row = last_row_check(src_header_row)
    
    
    'Find the last used rows & columns of the designated worksheet
    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Check if the source header = designated worksheet header, if yes, copy and paste it after the last used row
    Range("A" & header_row).Select
    If src_last_row >= src_data_row Then
        For i = 1 To last_col
            
            For j = src_start_col To src_last_col
                final_header_name = Workbooks(wkbk).Worksheets(wksheet).Cells(header_row, i).Value
                src_header_name = Workbooks(src_bk).Worksheets(src_sheet).Cells(src_header_row, j).Value
                
                If Trim(final_header_name) = Trim(src_header_name) Then
    
                    Workbooks(src_bk).Worksheets(src_sheet).Range(Workbooks(src_bk).Worksheets(src_sheet).Cells(src_data_row, j), Workbooks(src_bk).Worksheets(src_sheet).Cells(src_last_row, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(last_row + 1, i).PasteSpecial xlPasteValuesAndNumberFormats
                    
                    
                     'Give source remark,source from which path, files, worksheet base on text in Control Column J ->"Add fixed text/pw"
    
                    If add_srctext <> "" And add_srctext <> "NA" Then
                        add_srctext_details = Split(add_srctext, ";")
                        For k = 0 To UBound(add_srctext_details)
                            add_to_col = Range(Trim(Left(add_srctext_details(k), InStr(1, add_srctext_details(k), "|", vbTextCompare) - 1)) & "1").Column
                            add_value = Split(Trim(Right(add_srctext_details(k), Len(add_srctext_details(k)) - InStr(1, add_srctext_details(k), "|", vbTextCompare))), ",")
                    
                            Range(Cells(last_row + 1, add_to_col), Cells(last_row + src_last_row - src_data_row + 1, add_to_col)).Value = add_value
                        Next k
                    End If
                
                    
                    
                    Exit For
                    
                End If
            Next j
        Next i
    End If
    Range("A" & header_row).Select
    Application.CutCopyMode = False
    'If file path & file name are provided, close file
    If (src_path & src_bk) <> "" And src_bk <> ThisWorkbook.Name Then
        Call Close_File(src_path, src_bk, "UNCHANGE")
    End If
    
End Sub


Sub Refresh_Pivot(wkbk, wksheet)
'To refresh pivot tables in the excel file

    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If

    If (UCase(wksheet) = "ALL") Then
    'Refresh all pivot tables in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If check_ctrl = 0 Then
                For Each PT In ws.PivotTables
                    PT.SourceData = check_pivot_range(PT.SourceData)
                    PT.RefreshTable
                Next PT
            End If
        Next ws

    Else
    'Refresh pivot table on specific worksheet
        For Each PT In Workbooks(wkbk).Worksheets(wksheet).PivotTables
            PT.RefreshTable
        Next

    End If
    
End Sub


Sub Filter(wkbk, wksheet, header_row, filter_by)
'Filter based on user requested column & values

    'Use current workbook if user input is blank
    Dim i, j As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    
    'Find the last used header columns & rows
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Get user filter requirements
    filter_by_details = Split(filter_by, ";")
    
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    For i = 0 To UBound(filter_by_details)
        filter_by_col = Range(Trim(Left(filter_by_details(i), InStr(1, filter_by_details(i), "|", vbTextCompare) - 1)) & "1").Column
        filter_by_value = Split(Trim(Right(filter_by_details(i), Len(filter_by_details(i)) - InStr(1, filter_by_details(i), "|", vbTextCompare))), ",")
        
        'To check if filter criteria = blank
            For j = 0 To UBound(filter_by_value)
                If Trim(filter_by_value(j)) = "(blank)" Then
                    filter_by_value(j) = ""
                End If
            Next j
            
        'Apply filter
        
        
        If UBound(filter_by_value) = 0 And filter_by_value(0) = "<>" Then
            ActiveSheet.Rows(header_row & ":" & last_row).AutoFilter Field:=filter_by_col, Criteria1:="<>", Operator:=xlFilterValues
        Else
            ActiveSheet.Rows(header_row & ":" & last_row).AutoFilter Field:=filter_by_col, Criteria1:=filter_by_value, Operator:=xlFilterValues
        End If
        
    Next i
    
       
End Sub


Sub Sorting(wkbk, wksheet, header_row, sort_by)
'Sort based on user requested column & order

    'Use current workbook if user input is blank
    Dim i As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    
    'Find the last used header rows and last cell
    last_row = last_row_check(header_row)
    last_cell = last_cell_check(header_row)

    'Get user sorting requirements (which col, ordering by, custom list if applicable)
    ActiveSheet.Sort.SortFields.Clear
    sorting_details = Split(sort_by, ";")
    For i = 0 To UBound(sorting_details)
        
        sort_col = Trim(Left(sorting_details(i), InStr(1, sorting_details(i), "|", vbTextCompare) - 1))
        need_custom = InStr(1, sorting_details(i), "(", vbTextCompare)
        
        'Check if custom sorting is required or not And get ordering by (Ascending/Descending)
        If need_custom = 0 Then
            sort_order = Trim(Right(sorting_details(i), Len(sorting_details(i)) - InStr(1, sorting_details(i), "|", vbTextCompare)))
            
        Else
            sort_order = Mid(sorting_details(i), InStr(1, sorting_details(i), "|", vbTextCompare) + 1, need_custom - InStr(1, sorting_details(i), "|", vbTextCompare) - 1)
            sort_custom = Split(Mid(sorting_details(i), need_custom + 1, Len(sorting_details(i)) - need_custom - 1), ",")
            'Add custom sorting list
            Application.AddCustomList sort_custom
            sort_custom_count = Application.CustomListCount
            
        End If
        
        'Start sorting
        If UCase(sort_order) = "ASCENDING" Then
            ActiveSheet.Sort.SortFields.Add Key:=Range(sort_col & header_row + 1 & ":" & sort_col & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=sort_custom_count, DataOption:=xlSortNormal
        ElseIf UCase(sort_order) = "DESCENDING" Then
            ActiveSheet.Sort.SortFields.Add Key:=Range(sort_col & header_row + 1 & ":" & sort_col & last_row), SortOn:=xlSortOnValues, Order:=xlDescending, CustomOrder:=sort_custom_count, DataOption:=xlSortNormal
        End If
        
    Next i

    'Select Range and apply sort
    With ActiveSheet.Sort
        .SetRange Range("A" & header_row & ":" & last_cell)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


Sub Unfilter(wkbk, wksheet)
'Unfilter the data

    'Use current workbook if user input is blank
    Dim i As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    'Workbooks(wkbk).Worksheets(wksheet).Activate
        
    If UCase(wksheet) = "ALL" Then
    'Unfilter all worksheets in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If (ActiveSheet.AutoFilterMode And check_ctrl = 0) Then
                Set Rng = ActiveSheet.AutoFilter.Range
                If Rng.Rows.Count > Rng.SpecialCells(xlCellTypeVisible).Rows.Count Then
                    ActiveSheet.ShowAllData
                Else
                    ActiveSheet.AutoFilterMode = False
                    Rng.AutoFilter
                End If
            End If
            
        Next ws
    Else
    'Unfilter specified in workbooks
        Workbooks(wkbk).Worksheets(wksheet).Activate
        If ActiveSheet.AutoFilterMode Then
            Set Rng = ActiveSheet.AutoFilter.Range
            If Rng.Rows.Count > Rng.SpecialCells(xlCellTypeVisible).Rows.Count Then
                ActiveSheet.ShowAllData
            Else
                ActiveSheet.AutoFilterMode = False
                Rng.AutoFilter
            End If
        End If
    End If
    
    
End Sub


Sub add_text(wkbk, wksheet, header_row, start_row, text_to_add)
'To add fixed text to a filtered cell

    'Use current workbook if user input is blank
    Dim i As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
           
    'Find the last used header rows
    last_row = last_row_check(header_row)
    If start_row = "" Then start_row = header_row + 1
    If last_row = header_row Then Exit Sub
    
    'Get user text fill requirements (which col, and text)
    add_details = Split(text_to_add, ";")
    For i = 0 To UBound(add_details)
        to_add_col = Trim(Left(add_details(i), InStr(1, add_details(i), "|", vbTextCompare) - 1))
        to_add_text = Trim(Right(add_details(i), Len(add_details(i)) - InStr(1, add_details(i), "|", vbTextCompare)))
    
        'Only filled the text for the cell which is visible
        Range(to_add_col & header_row + 1).Select
        to_add_cell = Range(Selection, Range(to_add_col & last_row)).Select
        For Each to_add_cell In Selection
            If Not to_add_cell.Rows.Hidden Then
                to_add_cell.Value = to_add_text
            End If
        Next
    
    Next i
        
    
End Sub


Sub Delete_Col_Row(wkbk, wksheet, del_by)
'To delete specified rows and columns

    'Use current workbook if user input is blank
    Dim i, j As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
        
    'Get user delete requirements (column / Row and which col / row)
    del_detail = Split(del_by, ";")
    For i = 0 To UBound(del_detail)
        to_del = Trim(Left(del_detail(i), InStr(1, del_detail(i), "|", vbTextCompare) - 1))
        to_del_item = Split(Trim(Right(del_detail(i), Len(del_detail(i)) - InStr(1, del_detail(i), "|", vbTextCompare))), ",")
        
        'Adjust user input as a string
        to_del_items = ""
        For j = 0 To UBound(to_del_item)
            If InStr(1, to_del_item(j), ":", vbTextCompare) = 0 Then
                to_del_item(j) = Trim(to_del_item(j)) & ":" & Trim(to_del_item(j))
            End If
            
            If to_del_items = "" Then
                to_del_items = to_del_item(j)
            Else
                to_del_items = to_del_items & "," & to_del_item(j)
            End If
        
        Next j
        
        'Delete row/column
        If UCase(to_del) = "ROW" Then
            Range(to_del_items).EntireRow.Delete
        End If
        If UCase(to_del) = "COLUMN" Then
            Range(to_del_items).EntireColumn.Delete

        End If
    Next i

End Sub


Sub Copy_Sheet(src_path, src_bk, src_sheet, output_path, output_bk)
'Copy sheet base on control table information
    Application.DisplayAlerts = False
    Dim i As Integer
    'Check if source workbook name is blank
    If src_bk = "" Then
        src_bk = ActiveWorkbook.Name
    End If
    
    'Check if source path is blank
    If src_path = "" Then
        src_path = ActiveWorkbook.Path
    End If
        
    'Check if source path and output path end with "\"
    If Right(src_path, 1) <> "\" Then src_path = src_path & "\"
    If Right(output_path, 1) <> "\" Then output_path = output_path & "\"
        
    'Open the source workbook
    If workbook_is_open(src_bk) = False Then
        Call Open_File(src_path, src_bk, "READ")
    End If
    
    'Check if output workbook exists and open. If not, create the file
    to_del = False
    If file_exist(output_path & output_bk) Then
        Call Open_File(output_path, output_bk, "EDIT")
    Else
        Workbooks.Add
        ActiveWorkbook.Sheets(1).Name = "~"
        ActiveWorkbook.SaveAs output_path & output_bk
        to_del = True
    End If
    sht_name = ActiveSheet.Name
        
    'Split if input more than one more worksheet, and copy to output workbook
    to_copy_sheet = Split(src_sheet, "|")
    Workbooks(src_bk).Activate
    For i = 0 To UBound(to_copy_sheet)
        Workbooks(src_bk).Sheets(to_copy_sheet(i)).Copy Before:=Workbooks(output_bk).Sheets(sht_name)
    Next i
    
    'Delete the blank sheet for newly created workbook
    If to_del = True Then
        Workbooks(output_bk).Activate
        Worksheets("~").Delete
    End If
    
    Call Close_File(output_path, output_bk, "SAVE")
    
End Sub


Sub Delete_sheet(wk_bk, del_by)
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    'Split if input more than one more worksheet, and delete worksheet
    to_del_sheet = Split(del_by, "|")
    For i = 0 To UBound(to_del_sheet)
        Workbooks(wk_bk).Activate
            If sheet_exist(wk_bk, to_del_sheet(i)) = True Then
                Worksheets(to_del_sheet(i)).Delete
            End If
    Next i
        
End Sub

Sub Cut_File()
    Application.DisplayAlerts = False
    Dim i, j, k As Integer
    'Application.ScreenUpdating = False
    criteria_first_col = 9
    
    If wbThis Is Nothing Then
        Set wbThis = ThisWorkbook
    End If
    wbThis.Worksheets("Cut file").Activate
    cutFile_start = Range("B1").Value
    
    cutFile_no = Range(cutFile_start).End(xlDown).Row - Range(cutFile_start).Row
    cutFile_lastcol = Range(cutFile_start).End(xlToRight).Column
    
    'Loop Cut file Table data into an array
    ReDim cutfile(cutFile_no, cutFile_lastcol)
    For i = 1 To cutFile_no
        Range(cutFile_start).Offset(i, 0).Select
        For j = 1 To cutFile_lastcol
            cutfile(i, j) = Selection.Value
            Selection.Offset(0, 1).Select
        Next j
    Next i
    
    'Get the criteria header
    ReDim criteria(cutFile_lastcol - criteria_first_col)
    Range(cutFile_start).Offset(0, criteria_first_col).Select
    For i = 1 To cutFile_lastcol - criteria_first_col
        criteria(i) = Selection.Value
        Selection.Offset(0, 1).Select
    Next i
    
    For i = 1 To cutFile_no
        'Check file path format
        If Right(cutfile(i, 1), 1) <> "\" Then cutfile(i, 1) = cutfile(i, 1) & "\"
        If Right(cutfile(i, 5), 1) <> "\" Then cutfile(i, 5) = cutfile(i, 5) & "\"
        
        'Create file if file does not exist
        If file_exist(cutfile(i, 5) & cutfile(i, 6)) = False Then
            FileCopy cutfile(i, 1) & cutfile(i, 2), cutfile(i, 5) & cutfile(i, 6)
        End If
        
        'Open file and go to specific worksheet
        Call Open_File(cutfile(i, 5), cutfile(i, 6), "EDIT")
        current_ws = ActiveSheet.Name
        Workbooks(cutfile(i, 6)).Activate
        If cutfile(i, 3) = "" Then
            cutfile(i, 3) = ActiveSheet.Name
        End If
        Worksheets(cutfile(i, 3)).Activate
        
        'Remove filter and check the no. of column
        filter_flag = 0
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilterMode = False
            filter_flag = 1
        End If
        'lastcol = last_col_check(cutfile(i, 4))
                        
        'Start checking column with same name
        For j = 1 To UBound(criteria)
            If cutfile(i, j + criteria_first_col) <> "" Then
                'Insert blank row just below the header (avoid header is merged)
                Rows(cutfile(i, 4) + 1 & ":" & cutfile(i, 4) + 1).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Selection.AutoFilter
                
                Range("A" & cutfile(i, 4)).Select
                For k = 1 To last_col_check(cutfile(i, 4))
                    'Check if header = filter criteria header
                    If Selection.Value = criteria(j) Then
                        'Split if multiple filter value
                        Dim multiple_select As Variant
                        If InStr(cutfile(i, j + criteria_first_col), ";") Then
                            multiple_select = Split(cutfile(i, j + criteria_first_col), ";")
                               
                            For m = 0 To UBound(multiple_select)
                                multiple_select(m) = Trim(multiple_select(m))
                            Next m
                            'Apply filter for multiple values
                            Selection.AutoFilter Field:=Selection.Column, Criteria1:="<>" & multiple_select(0), Operator:=xlAnd, Criteria2:="<>" & multiple_select(1)
                        Else
                            'Apply filter for single value
                            Selection.AutoFilter Field:=Selection.Column, Criteria1:="<>" & cutfile(i, j + criteria_first_col)
                        End If
                        
                        'Delete irrelvant rows
                        Rows(cutfile(i, 4) + 1 & ":" & cutfile(i, 4) + 1).Select
                        Range(Selection, Selection.End(xlDown).End(xlDown)).Select
                        Selection.Delete Shift:=xlUp
                        
                        Exit For
                    End If
                    
                    Selection.Offset(0, 1).Select
                Next k
            End If
        Next j
        
        Range("A1").Select
        
        'To reset filter as source file
        If filter_flag = 1 Then
            Rows(cutfile(i, 4) & ":" & cutfile(i, 4)).Select
            Selection.AutoFilter
            Range("A1").Select
        End If
        
        'To delete column / row in the cut file
        If cutfile(i, 7) <> "" Then
            Call Delete_Col_Row(cutfile(i, 6), cutfile(i, 3), cutfile(i, 7))
            Range("A1").Select
        End If
        
        'To delete worksheet in the cut file
        If cutfile(i, 8) <> "" Then
            Call Delete_sheet(cutfile(i, 6), cutfile(i, 8))
        End If
        
        'To refresh pivot in the cut file
        If cutfile(i, 9) <> "" Then
            Call Refresh_Pivot(cutfile(i, 6), UCase(cutfile(i, 9)))
        End If
            
        If sheet_exist(cutfile(i, 6), current_ws) = True Then
            Range("A1").Select
        End If
        
        'Save and close the cut file
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Next i
    

End Sub


Sub List_file(src_path, src_type, wksheet)
'To list all the file name within the folder based on the file extension input
    
    'Check if source path end with "\"
    If Right(src_path, 1) <> "\" Then src_path = src_path & "\"
    

    '"Initializes" the dir function and returns the first file
    all_file = Dir(src_path & src_type)
    
    'Select worksheet for listing the folder and file name, check for last used row
    Worksheets(wksheet).Activate
    last_row = last_row_check(1)
    Range("A" & last_row).Select
    
    'Loop in the selected folder until no more files found
    Do While all_file <> ""
    
        'Check if file extension align with the input source type, print file name on worksheet (to cater *xls = *xlsx in vba)
        If Right(all_file, Len(src_type) - 1) = Right(src_type, Len(src_type) - 1) Or src_type = "*.*" Then
            'Print out the file path and file name
            Selection.Offset(1, 0).Select
            ActiveCell.Value = src_path
            ActiveCell.Offset(0, 1).Value = all_file
            
        End If
        
        'Calls next file subsequently
        all_file = Dir
        
    Loop

End Sub


Sub List_subfolder(src_path, wksheet)
'To list all the subfolders name within the folder
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(src_path)
    i = 1
    'loops through each file in the directory and prints their names and path
    Worksheets(wksheet).Activate
    last_row = last_row_check(1)
    Range("A" & last_row).Select
    
    For Each objSubFolder In objFolder.subfolders
    'print folder name
        Selection.Offset(1, 0).Select
        ActiveCell.Value = src_path
        ActiveCell.Offset(0, 1).Value = objSubFolder.Name
    Next objSubFolder
    
End Sub

Sub List_all_files_subfolders(src_path, wksheet)
'To list all the subfolders name, files name and files name in all subfolders within the folder
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(src_path)
    i = 1
    'loops through each file in the directory and prints their names and path
    Worksheets(wksheet).Activate
    last_row = last_row_check(1)
    Range("A" & last_row).Select
    
    For Each objSubFolder In objFolder.subfolders
    'print folder name
        Selection.Offset(1, 0).Select
        ActiveCell.Value = src_path
        ActiveCell.Offset(0, 1).Value = objSubFolder.Name
        'Call List_file(src_path & "\" & objSubFolder.Name, "*.*", wksheet)
        Call List_all_files_subfolders(src_path & "\" & objSubFolder.Name, wksheet)
    Next objSubFolder
    
    Call List_file(src_path, "*.*", wksheet)
    
End Sub


Sub Move_file(src_path, src_bk, output_path, move_type)
'To move file from one folder to another folder (allow to select overwrite or not)

    'Check if source path and output path end with "\"
    If Right(src_path, 1) <> "\" And src_path <> "" Then src_path = src_path & "\"
    If Right(output_path, 1) <> "\" And output_path <> "" Then output_path = output_path & "\"
    
    'Check if output folder does not exist, create the folder
    If folder_exist(output_path) = False Then
        MkDir (output_path)
    End If
        
    'Check if source file exists and file is not open
    If file_exist(src_path & src_bk) = True And workbook_is_open(src_bk) = False Then
        
        'Check if output folder has file with same name
        If file_exist(output_path & src_bk) = True Then
            'Check if user select "Overwrite"
            If move_type = "OVERWRITE" Then
                FileCopy src_path & src_bk, output_path & src_bk
                Kill src_path & src_bk
            End If
            
        Else
            'Move file to output folder
            Name src_path & src_bk As output_path & src_bk
            
        End If
    End If
        
End Sub


Sub Copy_file(src_path, src_bk, output_path, copy_type)
'To copy file from one folder to another folder (allow to select overwrite or not)

    'Check if source path and output path end with "\"
    If Right(src_path, 1) <> "\" And src_path <> "" Then src_path = src_path & "\"
    If Right(output_path, 1) <> "\" And output_path <> "" Then output_path = output_path & "\"
    
    'Check if output folder does not exist, create the folder
    If folder_exist(output_path) = False Then
        MkDir (output_path)
    End If
    
    'Check if source file exists
    If file_exist(src_path & src_bk) = True Then
    
        'Check if user select "Overwrite"
        If copy_type = "OVERWRITE" Then
            FileCopy src_path & src_bk, output_path & src_bk
            
        'Check if file with same name exists in output folder
        ElseIf file_exist(output_path & src_bk) = False Then
            FileCopy src_path & src_bk, output_path & src_bk
            
        End If
    End If
            
End Sub


Sub Copy_all_files_subfolders(fromPath, toPath)
'To copy all files and subfolders from one folder to another folder

    Dim FSO As Object

    If Right(fromPath, 1) = "\" Then fromPath = Left(fromPath, Len(fromPath) - 1)
    If Right(toPath, 1) = "\" Then toPath = Left(toPath, Len(toPath) - 1)

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(fromPath) = False Then
        MsgBox fromPath & " doesn't exist"
        Exit Sub
    End If

    toFolder = Mid(fromPath, InStrRev(fromPath, "\") + 1)
    MkDir (toPath & "\" & toFolder)
    FSO.CopyFolder Source:=fromPath, Destination:=toPath & "\" & toFolder
    
    
End Sub


Sub Rename_file(src_path, src_bk, output_bk)
'To rename file name

    'Check if source path and output path end with "\"
    If Right(src_path, 1) <> "\" And src_path <> "" Then src_path = src_path & "\"
        
    'Check if source file exists and file is not open
    If file_exist(src_path & src_bk) = True And workbook_is_open(src_bk) = False Then
        Name src_path & src_bk As src_path & output_bk

    End If
        
End Sub

Sub Delete_folder(src_path)
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    If Right(src_path, 1) = "\" Then
        src_path = Left(src_path, Len(src_path) - 1)
    End If

    If FSO.FolderExists(src_path) = False Then
        MsgBox src_path & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile src_path & "\*.*", True
    'Delete subfolders
    FSO.deletefolder src_path & "\*.*", True
    FSO.deletefolder src_path
    On Error GoTo 0

End Sub


Sub Delete_file(src_path, src_bk)

    If Right(src_path, 1) = "\" Then
        src_path = Left(src_path, Len(src_path) - 1)
    End If
    
    On Error Resume Next
    Kill src_path & "\" & src_bk
    On Error GoTo 0

    
End Sub


Sub DelBlankSheet(wkbk)
'To delete blank sheet

    'Application.DisplayAlerts = False
    Workbooks(wkbk).Activate
    For Each sheet In ActiveWorkbook.Sheets
        If Application.CountA(sheet.UsedRange.Cells) = 0 Then
            sheet.Delete
        End If
    Next
    'Application.DisplayAlerts = True

End Sub


Sub Paste_cell_as_value(wkbk, wksheet, paste_row, paste_col)
'To copy specific range and paste as value
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If UCase(wksheet) = "ALL" Then
    'Paste selected range as value for all worksheets  in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If (check_ctrl = 0) Then
                'Check to paste row
                paste_row_from = Left(paste_row, InStr(1, paste_row, ":") - 1)
                If UCase(Right(paste_row, Len(paste_row) - InStr(1, paste_row, ":"))) = "LAST" Then
                    paste_row_to = last_row_check(paste_row_from)
                Else
                    paste_row_to = Right(paste_row, Len(paste_row) - InStr(1, paste_row, ":"))
                End If
                
                'Check to paste column
                paste_col_from = Range(Left(paste_col, InStr(1, paste_col, ":") - 1) & "1").Column
                If UCase(Right(paste_col, Len(paste_col) - InStr(1, paste_col, ":"))) = "LAST" Then
                    paste_col_to = last_col_check(paste_row_from)
                Else
                    paste_col_to = Range(Right(paste_col, Len(paste_col) - InStr(1, paste_col, ":")) & "1").Column
                End If
                
                Range(Cells(paste_row_from, paste_col_from), Cells(paste_row_to, paste_col_to)).Select
                Selection.Copy
                ActiveCell.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
                Range("A1").Select
            End If
            
        Next ws
    Else
        'Check to paste row
        Workbooks(wkbk).Worksheets(wksheet).Activate
        paste_row_from = Left(paste_row, InStr(1, paste_row, ":") - 1)
        If UCase(Right(paste_row, Len(paste_row) - InStr(1, paste_row, ":"))) = "LAST" Then
            paste_row_to = last_row_check(paste_row_from)
        Else
            paste_row_to = Right(paste_row, Len(paste_row) - InStr(1, paste_row, ":"))
        End If
        
        'Check to paste column
        paste_col_from = Range(Left(paste_col, InStr(1, paste_col, ":") - 1) & "1").Column
        If UCase(Right(paste_col, Len(paste_col) - InStr(1, paste_col, ":"))) = "LAST" Then
            paste_col_to = last_col_check(paste_row_from)
        Else
            paste_col_to = Range(Right(paste_col, Len(paste_col) - InStr(1, paste_col, ":")) & "1").Column
        End If
        
        Range(Cells(paste_row_from, paste_col_from), Cells(paste_row_to, paste_col_to)).Select
        Selection.Copy
        ActiveCell.PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        Range("A1").Select
    End If


End Sub


Sub Paste_sheet_as_value(wkbk, wksheet)
'To copy the whole sheet and paste as value
    Dim i As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    If UCase(wksheet) = "ALL" Then
    'Paste all worksheets as value in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If (check_ctrl = 0) Then
                Cells.Select
                Selection.Copy
                ActiveCell.PasteSpecial xlPasteValues
                Range("A1").Select
            End If
            
        Next ws
    Else
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Cells.Select
        Selection.Copy
        ActiveCell.PasteSpecial xlPasteValues
        Range("A1").Select
    End If
    
    Application.CutCopyMode = False

End Sub


Sub Send_Email(header_row)
    Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    'criteria_first_col = 5
    Dim i, j As Integer
    If wbThis Is Nothing Then
        Set wbThis = ThisWorkbook
    End If
    wbThis.Worksheets("Email").Activate
    email_start = "A" & header_row
    
    email_no = Range(email_start).End(xlDown).Row - Range(email_start).Row
    email_lastcol = Range(email_start).End(xlToRight).Column
    
    'Loop Cut file Table data into an array
    ReDim emailToSend(email_no, email_lastcol)
    For i = 1 To email_no
        Range(email_start).Offset(i, 0).Select
        For j = 1 To email_lastcol
            emailToSend(i, j) = Selection.Value
            Selection.Offset(0, 1).Select
        Next j
    Next i
    
    'To send email
    For i = 1 To email_no
        'Set Outlook application object
        On Error GoTo ErrHandler
    
        Dim objOutlook As Object
        Set objOutlook = CreateObject("Outlook.Application")
    
        'Create Email Object
        Dim objEmail As Object
        Set objEmail = objOutlook.CreateItem(olMailItem)
        
        With objEmail
            .To = emailToSend(i, 1)
            .cc = emailToSend(i, 2)
            .bcc = emailToSend(i, 3)
            .Subject = emailToSend(i, 4)
            .Body = emailToSend(i, 5)
            '.Display        ' Display email only
            .Send
        End With
        
        ' CLEAR.
        Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:

    Next i


End Sub

Sub Expand_group(wkbk, wksheet)
'To show gourped columns / rows
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If

    If UCase(wksheet) = "ALL" Then
    'Unfilter all worksheets in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If check_ctrl = 0 Then
                ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
            End If
        Next ws
    Else
    'Unfilter specified in workbooks
        Workbooks(wkbk).Worksheets(wksheet).Activate
        ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
    End If

    
End Sub


Sub Collapse_group(wkbk, wksheet)
'To show gourped columns / rows
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    
    If UCase(wksheet) = "ALL" Then
    'Unfilter all worksheets in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If check_ctrl = 0 Then
                ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
            End If
        Next ws
    Else
    'Unfilter specified in workbooks
        Workbooks(wkbk).Worksheets(wksheet).Activate
        ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    End If

    
End Sub


Sub Protect_sheet(wkbk, wksheet, pw)
'To protect worksheet with Or without pw
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    
    If UCase(wksheet) = "ALL" Then
    'Unfilter all worksheets in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If check_ctrl = 0 Then
                ws.Protect Password:=pw
            End If
        Next ws
    Else
    'Unfilter specified in workbooks
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Worksheets(wksheet).Protect Password:=pw
    End If
End Sub


Sub Unprotect_sheet(wkbk, wksheet, pw)
'To unprotect worksheet with Or without pw
    Dim i As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    
    If UCase(wksheet) = "ALL" Then
    'Unfilter all worksheets in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            ws.Activate
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            If check_ctrl = 0 Then
                ws.Unprotect Password:=pw
            End If
        Next ws
    Else
    'Unfilter specified in workbooks
        Workbooks(wkbk).Worksheets(wksheet).Activate
        Worksheets(wksheet).Unprotect Password:=pw
    End If
End Sub


Sub Protect_workbook(wkbk, pw)
'To protect workbook with Or without pw
   
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    Workbooks(wkbk).Protect Password:=pw


End Sub


Sub Unprotect_workbook(wkbk, pw)
'To unprotect workbook with Or without pw
   
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    Workbooks(wkbk).Unprotect Password:=pw
    
End Sub


Sub Turn_on_auto_cal()

        Application.Calculation = xlAutomatic
        
        
End Sub

Sub Turn_off_auto_cal()

        Application.Calculation = xlManual
        
        
End Sub


Sub Change_dot_to_slash(wkbk, wksheet, header_row, start_row, dot_col)
'Change all dot to slash for specified column
    Dim i, j As Integer
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    If start_row = "" Then
        start_row = header_row + 1
    End If
    
    last_row = last_row_check(header_row)
    to_chg_col = Split(dot_col, "|")
    For i = 0 To UBound(to_chg_col)
        For j = start_row To last_row
            Range(to_chg_col(i) & j).Select
            Selection.Value = DateValue(Left(Selection.Value, 2) & "/" & Mid(Selection.Value, 4, 2) & "/" & Right(Selection.Value, 4))
        Next j
        Range(to_chg_col(i) & start_row & ":" & to_chg_col(i) & last_row).NumberFormat = "dd/mm/yyyy"
    Next i
    
End Sub


Function last_col_check(header_row) As Double
'Find the last used header column and return as number

    Dim maxcol As Long
    
    maxcol = 0
    For Each col In Columns
        If Application.WorksheetFunction.CountA(col) > 0 Then
            If col.Column > maxcol Then
                maxcol = col.Column
            End If
        End If
    Next
    
    last_col_check = maxcol

'    last_col_check = Cells(header_row, Columns.Count).End(xlToLeft).Column
    
End Function


Function last_row_check(header_row) As Double
'Find the last used row and return as number
    Dim maxrow As Long
    
    maxrow = 0
    For Each Row In Rows
        If Application.WorksheetFunction.CountA(Row) > 0 Then
            If Row.Row > maxrow Then
                maxrow = Row.Row
            End If
        End If
    Next
    
    last_row_check = maxrow
    
'    last_col = last_col_check(header_row)
'    Range("A" & header_row).Select
'    For col = 1 To last_col
'        If Cells(Rows.Count, col).End(xlUp).Row > last_row_check Then
'            last_row_check = Cells(Rows.Count, col).End(xlUp).Row
'        End If
'
'        Selection.Offset(0, 1).Select
'
'    Next col

End Function


Function last_cell_check(header_row) As String
'Check last cell based on last used column & row

    last_col_cell = Cells(header_row, Columns.Count).End(xlToLeft).Address
    last_row = last_row_check(header_row)
    last_cell_check = Left(last_col_cell, InStr(2, last_col_cell, "$", vbTextCompare) - 1) & last_row
    
End Function


Function cell_in_pivot() As Boolean
'Check if the selected cell belongs to a pivot

    Dim PT As PivotTable
    On Error Resume Next
    Set PT = ActiveCell.PivotTable
    On Error GoTo 0
    
    If PT Is Nothing Then
        cell_in_pivot = False
    Else
        cell_in_pivot = True
    End If
    
End Function


Function folder_exist(fullpath) As Boolean
'Returns TRUE if the folder exists
    If Not Dir(fullpath, vbDirectory) = vbNullString Then folder_exist = True
End Function


Function file_exist(fullpath) As Boolean
'Returns TRUE if the file exists
    file_exist = (Dir(fullpath) <> "")
End Function


Function sheet_exist(wb, ws) As Boolean
'Returns TRUE if the worksheet exists
    On Error Resume Next
    Workbooks(wb).Activate
    sheet_exist = (ActiveWorkbook.Sheets(ws).Index > 0)
End Function


Function workbook_is_open(wbname) As Boolean
'Returns TRUE if the workbook is opened

    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    
    If Err = 0 Then
        workbook_is_open = True
    Else
        workbook_is_open = False
    End If
    
End Function


Function check_pivot_range(ByVal pivot_range As String) As String
    'Use: 1) Check if the source data of the pivot file contains one or more rows
    '     2) If more than or equal to 1 row, return the same source range
    '     3) If header only, select one more row for pivot
    
    Dim mystr, str1, str2, final_str, sheet, Cell_1, Cell_2 As String
    Dim i, j As Integer
     
    
    sheet = Left(pivot_range, InStr(pivot_range, "!"))
    Cell_1 = Mid(pivot_range, InStr(pivot_range, "!") + 2, InStr(pivot_range, ":") - InStr(pivot_range, "!") - 1)
    Cell_2 = Right(pivot_range, Len(pivot_range) - InStr(pivot_range, ":") - 1)
     
    str1 = Left(Cell_1, InStr(Cell_1, "C") - 1)
    str2 = Left(Cell_2, InStr(Cell_2, "C") - 1)
     
    If str1 = str2 Then
        str2 = str2 + 1
        final_str = sheet & "R" & Cell_1 & "R" & str2 & "C" & Right(Cell_2, Len(Cell_2) - InStr(Cell_2, "C"))
        check_pivot_range = final_str
    Else
        check_pivot_range = pivot_range
    End If
    
        
End Function



Sub Filter_Period(wkbk, wksheet, header_row, filter_by)
'Filter based on user requested column & values
    Dim i, j As Integer
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    
    'Find the last used header columns & rows
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Get user filter requirements
    filter_by_details = Split(filter_by, ";")
    
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    For i = 0 To UBound(filter_by_details)
        filter_by_col = Range(Trim(Left(filter_by_details(i), InStr(1, filter_by_details(i), "|", vbTextCompare) - 1)) & "1").Column
        filter_by_value = Split(Trim(Right(filter_by_details(i), Len(filter_by_details(i)) - InStr(1, filter_by_details(i), "|", vbTextCompare))), ",")
        
        'To check if filter criteria = blank
            For j = 0 To UBound(filter_by_value)
                If Trim(filter_by_value(j)) = "(blank)" Then
                    filter_by_value(j) = ""
                End If
            Next j
            
        'Apply filter
        'Check if filter value is date or other criteria
        If IsDate(Right(filter_by_value(0), 10)) Or Left(filter_by_value(0), 1) = ">" Or Left(filter_by_value(0), 1) = "<" Then
            If UBound(filter_by_value) = 0 Then
                ActiveSheet.Rows(header_row & ":" & last_row).AutoFilter Field:=filter_by_col, Criteria1:=filter_by_value(0), Operator:=xlAnd
            Else
                ActiveSheet.Rows(header_row & ":" & last_row).AutoFilter Field:=filter_by_col, Criteria1:=filter_by_value(0), Operator:=xlAnd, Criteria2:=filter_by_value(1)
            End If
        Else
            ActiveSheet.Rows(header_row & ":" & last_row).AutoFilter Field:=filter_by_col, Criteria1:=filter_by_value, Operator:=xlFilterValues
        End If
            
    Next i
    
       
End Sub

Sub Append_All_notCloseFile(wkbk, wksheet, ind_row, header_row, add_srctext, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
'Using "Source_Start" as an indicator to append all source data to working worksheet
'Note: This function does not check the column name.
'      Please make sure the columns are same as the source data
    Dim i As Integer
    'Skip this step if user input N/A in source workbook
    If src_bk = "N/A" Then
       Exit Sub
    End If

    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    'Find the last used header columns & rows
    Call Unfilter(wkbk, wksheet)
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Find the indicator "Source_Start" in worksheet
    Range("A" & ind_row).Select
    For i = 1 To last_col
        If UCase(Selection.Value) = "SOURCE_START" Then
            Cells(last_row + 1, Selection.Column).Select
            
            'If file path & file name are provided, open file. If not, use current workbook
            If (src_path & src_bk) <> "" Then
                'If file has already open, it will not be re-opened
                If workbook_is_open(src_bk) = False Then
                    Call Open_File(src_path, src_bk, "READ")
                End If
                         
            Else
                src_bk = ThisWorkbook.Name
        
            End If
            
            'If no worksheet name, set current worksheet as source worksheet
            If src_sheet = "" Then
                src_sheet = Workbooks(src_bk).ActiveSheet.Name
            End If
            Workbooks(src_bk).Worksheets(src_sheet).Activate
            
            'Select data to be copied. This can select data with blank cells but separated by whole blank row
            If src_data_col = "" Then src_data_col = "A"
            src_last_cell = last_cell_check(src_header_row)
            src_last_row = last_row_check(src_header_row)

            If src_last_row >= src_data_row Then
                Range(Cells(src_data_row, src_data_col), src_last_cell).Select
                
                'Copy and paste as value
                Selection.Copy
                Workbooks(wkbk).Worksheets(wksheet).Activate
                ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
                Application.CutCopyMode = False
                
                 'Give source remark,source from which path, files, worksheet base on text in Control Column J ->"Add fixed text/pw"
    
                If add_srctext <> "" And add_srctext <> "NA" Then
                    add_srctext_details = Split(add_srctext, ";")
                    For k = 0 To UBound(add_srctext_details)
                        add_to_col = Range(Trim(Left(add_srctext_details(k), InStr(1, add_srctext_details(k), "|", vbTextCompare) - 1)) & "1").Column
                        add_value = Split(Trim(Right(add_srctext_details(k), Len(add_srctext_details(k)) - InStr(1, add_srctext_details(k), "|", vbTextCompare))), ",")
                    
                        Range(Cells(last_row + 1, add_to_col), Cells(last_row + src_last_row - src_data_row + 1, add_to_col)).Value = add_value
                    Next k
                End If
                
            End If
            
            Exit For
        End If
        
        Selection.Offset(0, 1).Select
        
    Next i
    Range("A" & ind_row).Select
End Sub

Sub Append_by_Col_Name_notCloseFile(wkbk, wksheet, header_row, add_srctext, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
'To check and copy columns with same name to designated worksheet
    Dim i, j As Integer
    'Skip this step if user input N/A in source workbook
    If src_bk = "N/A" Then
       Exit Sub
    End If
    
    
    'If file path & file name are provided, open file. If not, use current workbook
    If (src_path & src_bk) <> "" Then
        'If file has already open, it will not be re-opened
        If workbook_is_open(src_bk) = False Then
            Call Open_File(src_path, src_bk, "READ")
        End If
    Else
        src_bk = ThisWorkbook.Name

    End If
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    'If no worksheet name, set current worksheet as source worksheet
    If src_sheet = "" Then
        src_sheet = Workbooks(src_bk).ActiveSheet.Name
    End If
    Workbooks(src_bk).Worksheets(src_sheet).Activate
    
    'Find the last used header rows & columns of the source worksheet
    If src_data_col = "" Then src_data_col = "A"
    src_start_col = Range(src_data_col & src_header_row).Column
    src_last_col = last_col_check(src_header_row)
    src_last_row = last_row_check(src_header_row)
    
    
    'Find the last used rows & columns of the designated worksheet
    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
    'Check if the source header = designated worksheet header, if yes, copy and paste it after the last used row
    Range("A" & header_row).Select
    If src_last_row >= src_data_row Then
        For i = 1 To last_col
            
            For j = src_start_col To src_last_col
                final_header_name = Workbooks(wkbk).Worksheets(wksheet).Cells(header_row, i).Value
                src_header_name = Workbooks(src_bk).Worksheets(src_sheet).Cells(src_header_row, j).Value
                
                If Trim(final_header_name) = Trim(src_header_name) Then
    
                    Workbooks(src_bk).Worksheets(src_sheet).Range(Workbooks(src_bk).Worksheets(src_sheet).Cells(src_data_row, j), Workbooks(src_bk).Worksheets(src_sheet).Cells(src_last_row, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(last_row + 1, i).PasteSpecial xlPasteValuesAndNumberFormats
                    
                      'Give source remark,source from which path, files, worksheet base on text in Control Column J ->"Add fixed text/pw"
    
                    If add_srctext <> "" And add_srctext <> "NA" Then
                        add_srctext_details = Split(add_srctext, ";")
                        For k = 0 To UBound(add_srctext_details)
                            add_to_col = Range(Trim(Left(add_srctext_details(k), InStr(1, add_srctext_details(k), "|", vbTextCompare) - 1)) & "1").Column
                            add_value = Split(Trim(Right(add_srctext_details(k), Len(add_srctext_details(k)) - InStr(1, add_srctext_details(k), "|", vbTextCompare))), ",")
                    
                            Range(Cells(last_row + 1, add_to_col), Cells(last_row + src_last_row - src_data_row + 1, add_to_col)).Value = add_value
                        Next k
                    End If
                    
                    Exit For
                    
                End If
            Next j
        Next i
    End If
    Range("A" & header_row).Select
    Application.CutCopyMode = False

    
End Sub

Sub Change_text(wkbk, wksheet, header_row, start_row, text_to_change)
    Dim i As Integer
    Application.Calculation = xlAutomatic
'To add fixed text to a filtered cell

    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
           
'    'Find the last used header rows
'    last_row = last_row_check(header_row)
'    If start_row = "" Then start_row = header_row + 1
'    If last_row = header_row Then Exit Sub
    
    'Get user text fill requirements (which Range, and text)
    add_details = Split(text_to_change, ";")
    For i = 0 To UBound(add_details)
        to_add_range = Trim(Left(add_details(i), InStr(1, add_details(i), "|", vbTextCompare) - 1))
        to_add_text = Trim(Right(add_details(i), Len(add_details(i)) - InStr(1, add_details(i), "|", vbTextCompare)))
    
        'Only filled the text for the cell which is visible
        Range(to_add_range).Select
        For Each Rng In Selection
        
            If Not Rng.Rows.Hidden Then
                Rng.Value = to_add_text
            End If
        Next Rng
    Next i
        
    
End Sub

Sub Append_in_Same_Line(wkbk, wksheet, ind_row, header_row, start_row, src_path, src_bk, src_sheet, src_header_row, src_data_row, src_data_col)
    
    Application.DisplayAlerts = False
    Dim i, j As Integer
'To check and copy columns with same name to designated worksheet
    
    'Skip this step if user input N/A in source workbook
    If src_bk = "N/A" Then
       Exit Sub
    End If
    
    
    'If file path & file name are provided, open file. If not, use current workbook
    If (src_path & src_bk) <> "" Then
        'If file has already open, it will not be re-opened
        If workbook_is_open(src_bk) = False Then
            Call Open_File(src_path, src_bk, "READ")
        End If
    Else
        src_bk = ThisWorkbook.Name
        
    End If
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    'If no worksheet name, set current worksheet as source worksheet
    If src_sheet = "" Then
        src_sheet = Workbooks(src_bk).ActiveSheet.Name
    End If
    Workbooks(src_bk).Worksheets(src_sheet).Activate
    Call Unfilter(src_bk, src_sheet)
    
    'Find the last used header rows & columns of the source worksheet
    If src_data_col = "" Then src_data_col = "A"
    src_start_col = Range(src_data_col & src_header_row).Column
    src_last_col = last_col_check(src_header_row)
    src_last_row = last_row_check(src_header_row)
    
    
    'Find the last used rows & columns of the designated worksheet
    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)
    
   
    'Find the Start Column of the designated worksheet
    For i = 1 To last_col
        If UCase(Cells(ind_row, i)) = "SOURCE_START" Then
            Start_col = i
        ElseIf UCase(Cells(ind_row, i)) = "SOURCE_END" Then
            End_col = i
        End If
    Next i
    
     'Check if the source header = designated worksheet header, if yes, copy and paste it after the last used row
    Range("A" & header_row).Select
    If src_last_row >= src_data_row Then
        For i = Start_col To End_col
            
            For j = src_start_col To src_last_col
                final_header_name = Workbooks(wkbk).Worksheets(wksheet).Cells(header_row, i).Value
                src_header_name = Workbooks(src_bk).Worksheets(src_sheet).Cells(src_header_row, j).Value
                
                If Trim(final_header_name) = Trim(src_header_name) Then
    
                    Workbooks(src_bk).Worksheets(src_sheet).Range(Workbooks(src_bk).Worksheets(src_sheet).Cells(src_data_row, j), Workbooks(src_bk).Worksheets(src_sheet).Cells(src_last_row, j)).Copy
                    Workbooks(wkbk).Worksheets(wksheet).Cells(start_row, i).PasteSpecial xlPasteValuesAndNumberFormats
                    
                    Exit For
                    
                End If
            Next j
        Next i
    End If
    Range("A" & header_row).Select
    Application.CutCopyMode = False
    'If file path & file name are provided, close file
    If (src_path & src_bk) <> "" And src_bk <> ThisWorkbook.Name Then
        Call Close_File(src_path, src_bk, "UNCHANGE")
    End If
    
End Sub

Sub Remove_Duplicate(wkbk, wksheet, header_row, filter_by)
    'To Remove specific range duplicate records
    
    Dim i, j As Integer
    Dim myCol() As Variant
       
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    Workbooks(wkbk).Activate
    
    If wksheet = "" Then
        wksheet = ActiveSheet.Name
    End If
    Workbooks(wkbk).Worksheets(wksheet).Activate
    
    
    'Find the last used header columns & rows
    last_col = last_col_check(header_row)
    last_row = last_row_check(header_row)

    'Get user remove duplicate detail
    Removedup_col_start = Trim(Left(filter_by, InStr(1, filter_by, ":", vbTextCompare) - 1))
    Removedup_col_end = Trim(Right(filter_by, Len(filter_by) - InStr(1, filter_by, ":", vbTextCompare)))
    
    total_col = Range(Removedup_col_end & header_row).Column - Range(Removedup_col_start & header_row).Column + 1
    
    ReDim myCol(total_col - 1)
    
    For i = 0 To total_col - 1
        myCol(i) = i + 1
    Next i
    
      
    Text_Range = (Removedup_col_start & header_row) & ":" & (Removedup_col_end & last_row)
    
    Range(Text_Range).RemoveDuplicates Columns:=(myCol), Header:=xlYes

End Sub

Sub List_Worksheet(src_path, src_bk)
    Dim ws As Worksheet
    thisbk = ThisWorkbook.Name
    
    Workbooks(thisbk).Worksheets("File").Activate
    last_row = last_row_check(2)
    
    If last_row = 2 Then last_row = last_row + 1
    
    Range("A" & last_row).Value = src_path
    Range("B" & last_row).Value = src_bk
    
    For i = 2 To last_row - 1

     'Check if source path end with "\"
'        src_path = Cells(i + 1, 1).Value
'        src_bk = Cells(i + 1, 2).Value
     
        If Right(src_path, 1) <> "\" Then src_path = src_path & "\"
    
        If (src_path & src_bk) <> "" Then
                'If file has already open, it will not be re-opened
            If workbook_is_open(src_bk) = False Then
                Call Open_File(src_path, src_bk, "READ")
            End If
                         
            Else
                src_bk = ThisWorkbook.Name
        
        End If
    
        Workbooks(thisbk).Worksheets("File").Activate
        
        Range("B" & i + 1).Select
'        ActiveCell.Value = src_path
'        ActiveCell.Offset(0, 1).Value = src_bk
'        ActiveCell.Offset(0, 1).Select
    
    
        For Each ws In Workbooks(src_bk).Worksheets
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = ws.Name
        Next
    Next i
    
        
    
End Sub

Sub Pivot_Filter(wkbk, wksheet, filter_by)
'To refresh pivot tables in the excel file

    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim i, j, k As Single
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    Set ws = Worksheets(wksheet)
    
    'Get user filter requirements
    filter_by_details = Split(filter_by, ";")
    
    For i = 0 To UBound(filter_by_details)
        filter_by_field = Trim(Left(filter_by_details(i), InStr(1, filter_by_details(i), "|", vbTextCompare) - 1))
        filter_by_value = Split(Trim(Right(filter_by_details(i), Len(filter_by_details(i)) - InStr(1, filter_by_details(i), "|", vbTextCompare))), ",")
        
        
        'To check if filter criteria = blank
        For j = 0 To UBound(filter_by_value)
            If Trim(filter_by_value(j)) = "(blank)" Then
                filter_by_value(j) = ""
            End If
        Next j
            
        'Append Filter
        For Each PT In ws.PivotTables
            With ws.PivotTables(PT.Name).PivotFields(filter_by_field)
                .ClearAllFilters
                .CurrentPage = "(All)"

            'Unselect all Pivotitems
            For j = 1 To .PivotItems.Count - 1
                .PivotItems(j).Visible = False
            Next j
            
            'Select pivotitems is need
            For j = 1 To .PivotItems.Count - 1
                For k = 0 To UBound(filter_by_value)
                    If UCase(.PivotItems(j).Name) = UCase(filter_by_value(k)) Then
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

Sub Data_RefreshAll(wkbk, wksheet)

    Dim qry As WorkbookQuery
    Dim db As WorkbookConnection
    Dim PT As PivotTable
    Dim ws As Worksheet
    Dim i As Single
    
    
    'Use current workbook if user input is blank
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If

    If (UCase(wksheet) = "ALL") Then
    'Refresh all pivot tables in workbooks
        For Each ws In Workbooks(wkbk).Worksheets
            check_ctrl = 0
            
            'Check if the worksheet are control worksheet
            For i = 0 To UBound(ctrl_ws)
                If ws.Name = ctrl_ws(i) Then
                    check_ctrl = 1
                    
                End If
            Next i
            
            'Refresh All Pivot except Control,Email,Src,Cut file, file sheet
            If check_ctrl = 0 Then
                For Each PT In ws.PivotTables
                    PT.SourceData = check_pivot_range(PT.SourceData)
                    PT.RefreshTable
                Next PT
            End If
            
        Next ws

    Else
    'Refresh pivot table on specific worksheet
        For Each PT In Workbooks(wkbk).Worksheets(wksheet).PivotTables
            PT.RefreshTable
        Next
        
    End If
    
    i = wbThis.Worksheets("Query_LastRefresh").Range("A1").End(xlDown).Row
    If i = 1048576 Then
        i = 1
    End If
    
    
    
    For Each db In Workbooks(wkbk).Connections
        On Error Resume Next
        Workbooks(wkbk).Connections(db.Name).Refresh
        
        If Err.Number = 1004 Then
            last_refresh = "Can't access to data base"
        Else
            last_refresh = Now
        End If
        On Error GoTo 0
        wbThis.Worksheets("Query_LastRefresh").Range("A" & i + 1) = db.Name
        wbThis.Worksheets("Query_LastRefresh").Range("B" & i + 1) = last_refresh
        i = i + 1
    Next db
    
    

End Sub

Sub List_files(src_path, src_type, wksheet, ind_row, header_row)
'To list all the file name within the folder based on the file extension input
    
    'Check if source path end with "\"
    If Right(src_path, 1) <> "\" Then src_path = src_path & "\"
    If wksheet = "" Then wksheet = "File"
    

    '"Initializes" the dir function and returns the first file
    all_file = Dir(src_path & src_type)
    
    'Select worksheet for listing the folder and file name, check for last used row
    Worksheets(wksheet).Activate
    last_row = last_row_check(1)
    last_col = last_col_check(1)
    
    Range("A" & ind_row).Select
    
    'Locate the List File Columns
    For i = 1 To last_col
        Cells(ind_row, i).Select
        
        If UCase(Selection.Value) = "LIST_FILES" Then
            filepath_col = i
            filename_col = i + 1
            Cells(last_row, filepath_col).Select
            
    
    'Loop in the selected folder until no more files found
            Do While all_file <> ""
    
        'Check if file extension align with the input source type, print file name on worksheet (to cater *xls = *xlsx in vba)
                If UCase(Right(all_file, Len(src_type) - 1)) = UCase(Right(src_type, Len(src_type) - 1)) Or src_type = "*.*" Then
            'Print out the file path and file name
                    Selection.Offset(1, 0).Select
'                ActiveCell.Value = src_path
                    ActiveCell.Value = src_path
                    ActiveCell.Offset(0, 1).Value = all_file
            
                End If
        
        'Calls next file subsequently
                all_file = Dir
        
            Loop
        End If
    Next i
End Sub

Sub Copy_sheetdata(wkbk, wksheet, src_path, src_bk, src_sheet, src_header_row, src_data_col)

    Application.DisplayAlerts = False
    
'To check and copy columns with same name to designated worksheet
    
    'Skip this step if user input N/A in source workbook
    If src_bk = "N/A" Then
       Exit Sub
    End If
    
    
    'If file path & file name are provided, open file. If not, use current workbook
    If (src_path & src_bk) <> "" Then
        'If file has already open, it will not be re-opened
        If workbook_is_open(src_bk) = False Then
            Call Open_File(src_path, src_bk, "READ")
        End If
    Else
        src_bk = ThisWorkbook.Name
        
    End If
    
    If wkbk = "" Then
        wkbk = ThisWorkbook.Name
    End If
    
    'If no worksheet name, set current worksheet as source worksheet
    If src_sheet = "" Then
        src_sheet = Workbooks(src_bk).ActiveSheet.Name
    End If
    Workbooks(src_bk).Worksheets(src_sheet).Activate
    
    
    'Find the last used header rows & columns of the source worksheet
    If src_data_col = "" Then src_data_col = "A"
    src_start_col = Range("A1").Column
    src_last_col = last_col_check(src_header_row)
    src_last_row = last_row_check(src_header_row)
    
    
    'Find the last used rows & columns of the designated worksheet
    Workbooks(wkbk).Worksheets(wksheet).Activate
    Call Unfilter(wkbk, wksheet)
    
   
     'Check if the source header = designated worksheet header, if yes, copy and paste it after the last used row
    Range("A1").Select
    
    Workbooks(src_bk).Worksheets(src_sheet).Range(Workbooks(src_bk).Worksheets(src_sheet).Cells(1, 1), Workbooks(src_bk).Worksheets(src_sheet).Cells(src_last_row, src_last_col)).Copy
    Workbooks(wkbk).Worksheets(wksheet).Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
                    

    Range("A1").Select
    Application.CutCopyMode = False
    'If file path & file name are provided, close file
    If (src_path & src_bk) <> "" And src_bk <> ThisWorkbook.Name Then
        Call Close_File(src_path, src_bk, "UNCHANGE")
    End If

'    MsgBox "last Col" & src_last_col & "+++ Last Row" & src_last_row
End Sub
