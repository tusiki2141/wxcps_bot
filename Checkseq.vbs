Sub Checkseq()

Dim SDate01, EDate01, SDate06, EDate06, zSDate06, zEDate06 As Long
'Dim Hr01, row01, Nr06, row06  As Long
Dim Hr01, row01, Nr06, row06  As Double
Dim Hr06a As Double
Dim SID01, SID06, zSID06, Chk As String
Dim crow01, crow02 As Long
Dim start_time As Single

    Worksheets("IT2006").Select
    Range("A7").Select
    Selection.End(xlDown).Select
    row06 = Selection.Row
    
    Worksheets("IT2001").Select
    Range("A6").Select
    Selection.End(xlDown).Select
    row01 = Selection.Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    start_time = Timer
    'Get value from IT2001
    For i = 1 + 6 To row01
    'For i = 7191 To 7270
    
        Chk = Worksheets("IT2001").Cells(i, 26).Value
        
        If Chk = "" Then
             SID01 = Worksheets("IT2001").Cells(i, 1).Value
            SDate01 = Worksheets("IT2001").Cells(i, 7).Value
             EDate01 = Worksheets("IT2001").Cells(i, 8).Value
            Hr01 = Worksheets("IT2001").Cells(i, 19).Value
            crow01 = Worksheets("IT2001").Cells(i, 25).Value
            
            'Application.StatusBar = "Progress: " & i & " of " & row01 & ": " & Format(i / row01, "Percent")
            
            'Get value from IT2006
            Worksheets("IT2006").Select
            Range("A7").Select
            If SID06 = SID01 Then
                crow01 = crow02
            End If
            
            'MsgBox (i & " " & crow01)
            If crow01 <> 0 Then
            
                For j = crow01 To row06
                    SID06 = Cells(j, 1).Value
                    SDate06 = Cells(j, 17).Value
                    EDate06 = Cells(j, 18).Value
                    Hr06 = Cells(j, 16).Value
                    'MsgBox ("i: " & i & "j: " & j)
                    'MsgBox (SDate01 & " " & SDate06 & " " & EDate06 & " " & hr06 & " " & Cells(j, 21).Value & " " & Hr01)
                    If SID06 = SID01 And SDate01 >= SDate06 And SDate01 <= EDate06 And Hr06 - Cells(j, 21).Value - Hr01 >= 0 Then
                        Cells(j, 21).Value = Hr01 + Cells(j, 21).Value
                        Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2006").Cells(j, 20).Value
                        crow02 = j
                        Exit For
                    ElseIf SID06 = SID01 And SDate01 >= SDate06 And SDate01 <= EDate06 And Hr06 - Cells(j, 21).Value > 0 Then
                        Hr06a = (Hr06 - Cells(j, 21).Value - Hr01) * -1
                        Cells(j, 21).Value = Hr06
                        Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2006").Cells(j, 20).Value & " " & Hr06a
                        For z = j + 1 To j + 1
                            'MsgBox (hr06 & " " & Cells(j, 21).Value & " " & Hr01 & " " & i & " " & j & " " & Z & " " & Hr06a)
                            Hr06 = Cells(z, 16).Value
                            'MsgBox (hr06)
                            zSDate06 = Cells(z, 17).Value
                            zEDate06 = Cells(z, 18).Value
                            zSID06 = Cells(z, 1).Value
                            If zSID06 <> SID01 Then
                            Cells(j, 21).Value = Cells(j, 21).Value + Hr06a
                            Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2001").Cells(i, 26).Value & " " & "overtaken " & Hr06a
                            Else
                            If zSID06 = SID01 And SDate01 >= zSDate06 And SDate01 <= zEDate06 And Hr06 - Hr06a > 0 Then
                                Cells(z, 21).Value = Hr06a
                                Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2001").Cells(i, 26).Value & " " & Worksheets("IT2006").Cells(z, 20).Value & " " & Hr06a
                                crow02 = z
                                GoTo aa
                            Else
                                Cells(j, 21).Value = Cells(j, 21).Value + Hr06a
                                'MsgBox (Hr06a & " " & Cells(j, 21).Value)
                                Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2001").Cells(i, 26).Value & " " & "overtaken " & Hr06a
                                Worksheets("IT2006").Cells(j, 22).Value = "overtaken " & Hr06a
                            End If
                            End If
                        Next z
                    End If
                Next j
            End If
        End If
aa:
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
ActiveWorkbook.Save

    MsgBox (((Timer - start_time) / 60) & " mins")


    
End Sub
