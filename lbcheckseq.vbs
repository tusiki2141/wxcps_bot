Sub Checkseq()

'Dim SDate01, EDate01, SDate06, EDate06 As Long
Dim SDate01, EDate01, SDate06, EDate06 As Date
'Dim Hr01, row01, No06, row06, Ded06 As Long
Dim Hr01, row01, No06, row06, Ded06, Acc06, aa, bb As Double
Dim SID01, SID06, Chk As String
Dim crow01 As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Worksheets("IT2006").Select
    Range("A7").Select
    Selection.End(xlDown).Select
    row06 = Selection.Row
    
    Worksheets("IT2001").Select
    Range("A6").Select
    Selection.End(xlDown).Select
    row01 = Selection.Row
    'row01 = 97404
    'row06 = 40660
    
    'Get value from IT2001
    'For i = 97401 To row01
    For i = 1 + 6 To row01
        Chk = Worksheets("IT2001").Cells(i, 26).Value
        
        If Chk = "" Then
            SID01 = Worksheets("IT2001").Cells(i, 1).Value
            SDate01 = Worksheets("IT2001").Cells(i, 7).Value
            EDate01 = Worksheets("IT2001").Cells(i, 8).Value
            Hr01 = Worksheets("IT2001").Cells(i, 18).Value
            crow01 = Worksheets("IT2001").Cells(i, 25).Value
            
            'MsgBox (SID01 & " " & SDate01 & " " & EDate01 & " " & Hr01 & " " & crow01)
            'Application.StatusBar = "Progress: " & i & " of " & row01 & ": " & Format(i / row01, "Percent")
            
            'Get value from IT2006
            Worksheets("IT2006").Select
            Range("A7").Select
            If crow01 <> 0 Then
            
                For j = crow01 To row06
                    SID06 = Cells(j, 1).Value
                    SDate06 = Cells(j, 17).Value
                    EDate06 = Cells(j, 18).Value
                    Hr06 = Cells(j, 16).Value
                    Ded06 = Cells(j, 19).Value
                    Acc06 = Cells(j, 21).Value
                    'Ded06 = Cells(j, 16).Value
                    
                    'MsgBox (SID06 & " " & SDate06 & " " & EDate06 & " " & Hr06 & " " & Ded06)
                    'If SID06 = SID01 And SDate01 >= SDate06 And SDate01 <= EDate06 And Cells(j, 21).Value < Ded06 Then
                    If SID06 = SID01 And SDate01 >= SDate06 And SDate01 <= EDate06 And Acc06 < Ded06 Then
                        If Hr01 + Cells(j, 21).Value >= Ded06 Then
                                aa = Cells(j + 1, 19).Value
                                bb = Hr01 + Cells(j, 21).Value - Ded06
                            'If  Cells(j + 1, 21).Value > = Hr01 + Cells(j, 21).Value - Ded06 Then
                            If Round(aa, 2) >= Round(bb, 2) Then
                                Cells(j + 1, 21).Value = Hr01 + Cells(j, 21).Value - Ded06
                                Cells(j, 21).Value = Ded06
                                Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2006").Cells(j, 20).Value & " " & Worksheets("IT2006").Cells(j + 1, 20).Value
                                Exit For
                            Else
                                Cells(j + 2, 21).Value = Hr01 + Cells(j, 21).Value - Ded06
                                Cells(j, 21).Value = Ded06
                                Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2006").Cells(j, 20).Value & " " & Worksheets("IT2006").Cells(j + 2, 20).Value
                                Exit For
                            End If
                        Else
                            Cells(j, 21).Value = Hr01 + Cells(j, 21).Value
                            Worksheets("IT2001").Cells(i, 26).Value = Worksheets("IT2006").Cells(j, 20).Value
                            
                        Exit For
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox ("Complete")
    
End Sub

