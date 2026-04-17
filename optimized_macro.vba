Option Explicit

' =====================================================
' ÓÅ»¯ºóµÄ¼ì²éÐòÁÐºê
' ¹¦ÄÜ£ºÔÚIT2001ºÍIT2006¹¤×÷±íÖ®¼ä½øÐÐ¹¤Ê±Æ¥Åä¼ÆËã
' ÓÅ»¯ÖØµã£º
' 1. Ïû³ý²»±ØÒªµÄÑ¡Ôñ²Ù×÷
' 2. Ê¹ÓÃÊý×éÅúÁ¿´¦ÀíÊý¾Ý
' 3. ¸Ä½ø±äÁ¿ÃüÃûºÍ´úÂë½á¹¹
' 4. ÓÅ»¯Ñ­»·Âß¼­
' =====================================================

Sub CheckseqOptimized()
    
    ' ÉùÃ÷±äÁ¿
    Dim startTime As Single
    Dim ws2001 As Worksheet, ws2006 As Worksheet
    Dim lastRow2001 As Long, lastRow2006 As Long
    Dim data2001() As Variant, data2006() As Variant
    Dim i As Long, j As Long, z As Long
    Dim currentRow2001 As Long, currentRow2006 As Long
    Dim matchFound As Boolean
    
    ' ¼ÇÂ¼¿ªÊ¼Ê±¼ä
    startTime = Timer
    
    ' ÉèÖÃ¹¤×÷±íÒýÓÃ
    Set ws2001 = Worksheets("IT2001")
    Set ws2006 = Worksheets("IT2006")
    
    ' ¹Ø±ÕÆÁÄ»¸üÐÂºÍ×Ô¶¯¼ÆËãÒÔÌá¸ßÐÔÄÜ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' »ñÈ¡Êý¾Ý·¶Î§
    lastRow2001 = ws2001.Cells(ws2001.Rows.Count, "A").End(xlUp).Row
    lastRow2006 = ws2006.Cells(ws2006.Rows.Count, "A").End(xlUp).Row
    
    ' ¼ì²éÊÇ·ñÓÐ×ã¹»µÄÊý¾Ý
    If lastRow2001 < 7 Or lastRow2006 < 7 Then
        MsgBox "Êý¾Ý²»×ã£¬Çë¼ì²é¹¤×÷±íÊý¾Ý"
        Exit Sub
    End If
    
    ' ½«Êý¾Ý¼ÓÔØµ½Êý×éÖÐ½øÐÐÅúÁ¿´¦Àí
    data2001 = ws2001.Range("A6:Z" & lastRow2001).Value
    data2006 = ws2006.Range("A7:V" & lastRow2006).Value
    
    ' Ö÷´¦ÀíÑ­»· - ±éÀúIT2001Êý¾Ý
    For i = 1 To UBound(data2001, 1)
        
        ' ¼ì²éÊÇ·ñÐèÒª´¦Àí£¨µÚ26ÁÐÎª¿Õ£©
        If IsEmpty(data2001(i, 26)) Or data2001(i, 26) = "" Then
            
            ' »ñÈ¡µ±Ç°ÐÐÊý¾Ý
            Dim employeeId As String
            Dim startDate As Long, endDate As Long
            Dim hours2001 As Double
            
            employeeId = data2001(i, 1)
            startDate = data2001(i, 7)
            endDate = data2001(i, 8)
            hours2001 = data2001(i, 19)
            currentRow2006 = data2001(i, 25) - 6 ' ÉÏ´ÎÆ¥ÅäµÄÎ»ÖÃ
            
            ' ¸üÐÂ×´Ì¬À¸ÏÔÊ¾½ø¶È
            Application.StatusBar = "½ø¶È: " & i & " / " & UBound(data2001, 1) & ": " & Format(i / UBound(data2001, 1), "Percent")
            
            ' Èç¹ûµ±Ç°ÐÐÓÐÉÏ´ÎÆ¥Åä¼ÇÂ¼£¬´Ó¸ÃÎ»ÖÃ¿ªÊ¼ËÑË÷
            If currentRow2006 = -6 Then currentRow2006 = UBound(data2006, 1)
           
            
            matchFound = False
            
            ' ÔÚIT2006ÖÐËÑË÷Æ¥ÅäÏî
            For j = currentRow2006 To UBound(data2006, 1)
                
                ' ¼ì²éÆ¥ÅäÌõ¼þ
                If data2006(j, 1) = employeeId And _
                   startDate >= data2006(j, 17) And _
                   startDate <= data2006(j, 18) Then
                    
                    Dim availableHours As Double, usedHours As Double
                    availableHours = data2006(j, 16)
                    usedHours = data2006(j, 21)
                    
                    ' Çé¿ö1£ºÓÐ×ã¹»¹¤Ê±
                    If availableHours - usedHours - hours2001 >= 0 Then
                        data2006(j, 21) = usedHours + hours2001
                        data2001(i, 26) = data2006(j, 20)
                        data2001(i, 25) = j ' ¼ÇÂ¼Æ¥ÅäÎ»ÖÃ
                        matchFound = True
                        Exit For
                        
                    ' Çé¿ö2£º¹¤Ê±²»×ã£¬ÐèÒªÒç³ö´¦Àí
                    ElseIf availableHours - usedHours > 0 Then
                        Dim overflowHours As Double
                        overflowHours = (availableHours - usedHours - hours2001) * -1
                        
                        ' ·ÖÅäµ±Ç°¿ÉÓÃµÄ¹¤Ê±
                        data2006(j, 21) = availableHours
                        data2001(i, 26) = data2006(j, 20) & " " & overflowHours
                        
                        ' ´¦ÀíÒç³öµ½ÏÂÒ»ÐÐ
                        matchFound = HandleHourOverflow(data2006, data2001, i, j, employeeId, startDate, overflowHours)
                        Exit For
                    End If
                End If
            Next j
            
            ' Èç¹ûÃ»ÓÐÕÒµ½Æ¥ÅäÏî£¬¼ÇÂ¼×´Ì¬
            If Not matchFound Then
                data2001(i, 26) = "Î´ÕÒµ½Æ¥ÅäÏî"
            End If
        End If
    Next i
    
    ' ½«´¦ÀíºóµÄÊý¾ÝÐ´»Ø¹¤×÷±í
    ws2001.Range("A6:Z" & lastRow2001).Value = data2001
    ws2006.Range("A7:V" & lastRow2006).Value = data2006
    
    ' »Ö¸´Ó¦ÓÃ³ÌÐòÉèÖÃ
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    ' ±£´æ¹¤×÷²¾
    ActiveWorkbook.Save
    
    ' ÏÔÊ¾´¦ÀíÊ±¼ä
    Dim elapsedTime As Single
    elapsedTime = Timer - startTime
    MsgBox "´¦ÀíÍê³É£¡ºÄÊ±: " & Format(elapsedTime / 60, "0.00") & " ·ÖÖÓ"
    
End Sub

' =====================================================
' ´¦Àí¹¤Ê±Òç³öµ½ÏÂÒ»ÐÐµÄº¯Êý
' =====================================================
Private Function HandleHourOverflow(ByRef data2006() As Variant, ByRef data2001() As Variant, _
                                   ByVal row2001 As Long, ByVal startRow2006 As Long, _
                                   ByVal employeeId As String, ByVal startDate As Long, _
                                   ByVal overflowHours As Double) As Boolean
    
    Dim result As Boolean
    result = False
    
    ' ¼ì²éÏÂÒ»ÐÐÊÇ·ñ¿ÉÒÔ½ÓÊÕÒç³ö¹¤Ê±
    If startRow2006 + 1 <= UBound(data2006, 1) Then
        
        Dim nextRow As Long
        nextRow = startRow2006 + 1
        
        ' ¼ì²éÏÂÒ»ÐÐÊÇ·ñÂú×ãÌõ¼þ
        If data2006(nextRow, 1) = employeeId And _
           startDate >= data2006(nextRow, 17) And _
           startDate <= data2006(nextRow, 18) And _
           data2006(nextRow, 16) - overflowHours > 0 Then
            
            ' ·ÖÅä¹¤Ê±µ½ÏÂÒ»ÐÐ
            data2006(nextRow, 21) = overflowHours
            data2001(row2001, 26) = data2001(row2001, 26) & " " & data2006(nextRow, 20) & " " & overflowHours
            data2001(row2001, 25) = nextRow
            result = True
            
        Else
            ' ÎÞ·¨·ÖÅä£¬±ê¼ÇÎªÒç³ö
            data2006(startRow2006, 22) = "Òç³ö " & overflowHours
            data2001(row2001, 26) = data2001(row2001, 26) & " Òç³ö " & overflowHours
            data2006(startRow2006, 21) = data2006(startRow2006, 21) + overflowHours
            result = False
        End If
    Else
        ' Ã»ÓÐÏÂÒ»ÐÐ£¬±ê¼ÇÎªÒç³ö
        data2006(startRow2006, 22) = "Òç³ö " & overflowHours
        data2001(row2001, 26) = data2001(row2001, 26) & " Òç³ö " & overflowHours
        data2006(startRow2006, 21) = data2006(startRow2006, 21) + overflowHours
        result = False
    End If
    
    HandleHourOverflow = result
    
End Function

' =====================================================
' ¸¨Öúº¯Êý£ºÇå³ýÖ®Ç°µÄ´¦Àí½á¹û
' =====================================================
Sub ClearPreviousResults()
    
    Dim ws2001 As Worksheet, ws2006 As Worksheet
    Set ws2001 = Worksheets("IT2001")
    Set ws2006 = Worksheets("IT2006")
    
    ' Çå³ýIT2001µÄµÚ26ÁÐ£¨´¦Àí½á¹û£©
    With ws2001
        If .Cells(.Rows.Count, "Z").End(xlUp).Row >= 6 Then
            .Range("Z6:Z" & .Cells(.Rows.Count, "Z").End(xlUp).Row).ClearContents
        End If
    End With
    
    ' Çå³ýIT2006µÄµÚ21-22ÁÐ£¨ÒÑÓÃ¹¤Ê±ºÍÒç³ö±ê¼Ç£©
    With ws2006
        If .Cells(.Rows.Count, "U").End(xlUp).Row >= 7 Then
            .Range("U7:V" & .Cells(.Rows.Count, "U").End(xlUp).Row).ClearContents
        End If
    End With
    
    MsgBox "ÒÑÇå³ýÖ®Ç°µÄ´¦Àí½á¹û"
    
End Sub

