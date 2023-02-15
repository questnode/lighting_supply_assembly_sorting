Sub Macro1()
'
' Macro1 Macro
'

'
    Dim col, row As Integer
    Dim code_all, code_family, code_replace As String
    
    col = 1
    row = 2
    
    Worksheets("Transform").Activate
    Cells(row, col).Select
    code_all = ActiveCell.Value
    code_family = Left(code_all, 2)
    
    If code_family = "AB" Then
        Debug.Print Left(code_all, Len(code_all) - 3)
        Debug.Print Mid(code_all, Len(code_all) - 3 + 1, 1)
        Debug.Print Right(code_all, 3 - 1)
        'code_all = Left(code_allRule, Len(code_all) - 3) & Mid(code_all, Len(code_all) - 2, 1) & Right(code_all, 3 - 1)
        code_replace = Mid(code_all, Len(code_all) - 3 + 1, 1)
        code_replace = Replace(Replace(Replace(code_replace, "2", "27"), "3", "30"), "4", "40")
        code_all = Left(code_all, Len(code_all) - 3) & code_replace & Right(code_all, 3 - 1)
        ActiveCell.Value = code_all
    End If
    
'    Cells(row + 1, col).Select
'    ActiveCell.Value = code
    
    
'    ActiveCell.Replace What:="AB11129", Replacement:="AB111279", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False

End Sub

Sub List_All_Family()
    
    Dim family_code, code_example, worksheet_read, worksheet_write As String
    Dim col_read, row_read, col_write, row_write As Integer
    
    worksheet_read = "PDF"
    worksheet_write = "Family_List"
    col_read = 1
    row_read = 2
    col_write = 2
    row_write = 227
    
'    Worksheets("Transform").Activate
    
    family_code = "Start"
    
    While family_code <> ""
        family_code = Left(Worksheets(worksheet_read).Cells(row_read, col_read), 3)
        If family_code <> Worksheets(worksheet_write).Cells(row_write - 1, col_write).Value Then
            Worksheets(worksheet_write).Cells(row_write, col_write).Value = family_code
            Worksheets(worksheet_write).Cells(row_write, col_write + 1).Value = Worksheets(worksheet_read).Cells(row_read, col_read).Value
            Worksheets(worksheet_write).Cells(row_write, col_write + 2).Value = Worksheets(worksheet_read).Cells(row_read, col_read + 1).Value
            row_write = row_write + 1
        End If
        row_read = row_read + 1
    Wend
    

End Sub

Sub Transform()
'
' Macro1 Macro
'

'
    Application.ScreenUpdating = False

    
    Dim col, row, rule_row, rule_group, replace_index As Integer
    Dim code_all, code_family, code_replace, worksheet_write, code_keep_left, code_keep_right As String
    
    worksheet_write = "New_Products_converted"
      
    col = 1
    row = 2
    
    code_family = "START"
    
    While code_family <> ""
        
        'Check if it's still the same family, if not, find new rule_row
        If code_family <> Left(Worksheets(worksheet_write).Cells(row, col).Value, 3) Then
            code_family = Left(Worksheets(worksheet_write).Cells(row, col).Value, 3)
            rule_row = Search_Rule(code_family)
        End If
        

        code_all = Worksheets(worksheet_write).Cells(row, col).Value
        
        If Worksheets("Family_List").Cells(rule_row, 5) = "NA" Then
            Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
        ElseIf Worksheets("Family_List").Cells(rule_row, 5) > Len(code_all) Then
            Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
        Else
            replace_index = Worksheets("Family_List").Cells(rule_row, 5)
            code_keep_left = Left(code_all, replace_index - 1)
            code_keep_right = Mid(code_all, replace_index + 1, Len(code_all) - replace_index)
            code_replace = Mid(code_all, replace_index, 1)

            If Worksheets("Family_List").Cells(rule_row, 4) = 1 Then
                If code_replace = 2 Then
                    code_replace = 27
                ElseIf code_replace = 3 Then
                    code_replace = 30
                ElseIf code_replace = 4 Then
                    code_replace = 40
                End If
            ElseIf Worksheets("Family_List").Cells(rule_row, 4) = 2 Then
                If code_replace = 0 Then
                    code_replace = 27
                ElseIf code_replace = 1 Then
                    code_replace = 30
                ElseIf code_replace = 2 Then
                    code_replace = 42
                ElseIf code_replace = 3 Then
                    code_replace = 60
                ElseIf code_replace = 6 Then
                    code_replace = "BL"
                 ElseIf code_replace = 7 Then
                    code_replace = "GR"
                ElseIf code_replace = 8 Then
                    code_replace = "RD"
                End If
            ElseIf Worksheets("Family_List").Cells(rule_row, 4) = 3 Then
                If code_replace = 20 Then
                    code_replace = 27
                ElseIf code_replace = 21 Then
                    code_replace = 30
                ElseIf code_replace = 22 Then
                    code_replace = 42
                ElseIf code_replace = 23 Then
                    code_replace = 60
                ElseIf code_replace = 26 Then
                    code_replace = "BL"
                 ElseIf code_replace = 27 Then
                    code_replace = "GR"
                ElseIf code_replace = 28 Then
                    code_replace = "RD"
                End If
            ElseIf Worksheets("Family_List").Cells(rule_row, 4) = "NA" Then
                Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
            End If
            code_all = code_keep_left & code_replace & code_keep_right
            Worksheets(worksheet_write).Cells(row, col).Value = code_all
        End If
        
'        If row = 8697 Then
'            row = row
'        End If
        
        row = row + 1
    Wend
    
    Application.ScreenUpdating = True
    
'    Cells(row + 1, col).Select
'    ActiveCell.Value = code
    
    
'    ActiveCell.Replace What:="AB11129", Replacement:="AB111279", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False

End Sub

Sub Transform_Exceptions()
'
' Macro1 Macro
'

'
    Application.ScreenUpdating = False

    
    Dim col, row, row_end, rule_row, rule_group, replace_index As Integer
    Dim code_all, code_family, code_replace, worksheet_write, code_keep_left, code_keep_right As String
    
    worksheet_write = "Foglio1convert"
      
    col = 1
    row = 2
    row_end = 17626
    
    code_family = "START"
    
    While row < row_end
        
        'Check if it's still the same family, if not, find new rule_row
        If code_family <> Left(Worksheets(worksheet_write).Cells(row, col).Value, 3) Then
            code_family = Left(Worksheets(worksheet_write).Cells(row, col).Value, 3)
            rule_row = Search_Rule(code_family)
        End If
        

        code_all = Worksheets(worksheet_write).Cells(row, col).Value
        
        If Worksheets("Family_List").Cells(rule_row, 5) = "NA" Then
            Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
        ElseIf Worksheets("Family_List").Cells(rule_row, 5) > Len(code_all) Then
            Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
        Else
            replace_index = Worksheets("Family_List").Cells(rule_row, 5)
            code_keep_left = Left(code_all, replace_index - 1)
'            code_keep_right = Mid(code_all, replace_index + 1, Len(code_all) - replace_index)
            code_replace = Mid(code_all, replace_index, 2)

            If Worksheets("Family_List").Cells(rule_row, 4) = 9 Then
                If code_replace = 2 Then
                    code_replace = 27
                ElseIf code_replace = 3 Then
                    code_replace = 30
                ElseIf code_replace = 4 Then
                    code_replace = 40
                End If
            ElseIf Worksheets("Family_List").Cells(rule_row, 4) = 3 Then
                If code_replace = 20 Then
                    code_replace = 27
                ElseIf code_replace = 21 Then
                    code_replace = 30
                ElseIf code_replace = 22 Then
                    code_replace = 42
                ElseIf code_replace = 23 Then
                    code_replace = 60
                ElseIf code_replace = 26 Then
                    code_replace = "BL"
                 ElseIf code_replace = 27 Then
                    code_replace = "GR"
                ElseIf code_replace = 28 Then
                    code_replace = "RD"
                End If
            ElseIf Worksheets("Family_List").Cells(rule_row, 4) = "NA" Then
                Worksheets(worksheet_write).Cells(row, col + 7).Value = "No Change"
            End If
'            code_all = code_keep_left & code_replace & code_keep_right
            code_all = code_keep_left & code_replace
            Worksheets(worksheet_write).Cells(row, col).Value = code_all
        End If
        
'        If row = 8697 Then
'            row = row
'        End If
        
        row = row + 1
    Wend
    
    Application.ScreenUpdating = True
    
'    Cells(row + 1, col).Select
'    ActiveCell.Value = code
    
    
'    ActiveCell.Replace What:="AB11129", Replacement:="AB111279", LookAt:= _
'        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False

End Sub
Sub Remove_Items()

    Application.ScreenUpdating = False

    Dim col, row, row_write As Integer
'    Dim description, check_blank As String
    
    worksheet_read = "New_Products_converted"
    worksheet_write = "New_Products_cleaned"
      
    col = 3
    row = 2
    
    row_write = 2
    
'    description = "START"
    
    While Worksheets(worksheet_read).Cells(row, col).Value <> ""
'        description = Worksheets(worksheet_read).Cells(row, col).Value
        
        If InStr(Worksheets(worksheet_read).Cells(row, col).Value, "6000K") > 0 Or InStr(Worksheets(worksheet_read).Cells(row, col).Value, "CRI85") > 0 Then
            Worksheets(worksheet_read).Rows(row).Cut
            Worksheets(worksheet_write).Rows(row_write).Insert Shift:=xlDown
            If Worksheets(worksheet_read).Cells(row, col).Value = "" Then
                Worksheets(worksheet_read).Rows(row).Delete Shift:=xlUp
            End If
            row_write = row_write + 1
            row = row - 1
        End If

        row = row + 1
    Wend

    Application.ScreenUpdating = True

End Sub

Public Function Search_Rule(code_family) As Integer

    Dim rule_row, rule_col As Integer
    Dim seek_result As String
    
    rule_row = 2
    rule_col = 1
    seek_result = "START"
    
    Do While seek_result <> ""
        seek_result = Worksheets("Family_List").Cells(rule_row, rule_col).Value
        If seek_result = code_family Then
            Exit Do
        Else
            rule_row = rule_row + 1
        End If
    Loop

    Search_Rule = rule_row
    
End Function

Sub Find_Price()

    Application.ScreenUpdating = False
    
    Dim col_write, row_write, col_read, row_read, price_row As Integer
    Dim code_all, worksheet_write, worksheet_read, seek_result, price As String
    
    worksheet_read = "PDF"
    worksheet_write = "Transform"
      
    col_write = 2
    row_write = 2
    col_read = 2
    row_read = 2
    
    code_all = "START"
    seek_result = "START"
    
    While code_all <> ""
        If code_all <> Worksheets(worksheet_write).Cells(row_write, col_write) Then
            code_all = Worksheets(worksheet_write).Cells(row_write, col_write)

            Do While seek_result <> ""
                seek_result = Worksheets(worksheet_read).Cells(row_read, col_read)
                If seek_result = code_all Then
                    price = Worksheets(worksheet_read).Cells(row_read, col_read + 2)
                    Exit Do
                Else
                    price = "NA"
                    row_read = row_read + 1
                End If
            Loop
            row_read = 2
            seek_result = "START"
        End If
        
        Worksheets(worksheet_write).Cells(row_write, col_write + 6) = price
        If (row_write Mod 10000) = 0 Then
            row_write = row_write
        End If
        
        row_write = row_write + 1
    Wend

    Application.ScreenUpdating = True
End Sub
Sub Determine_Rules()

    Dim col, row, rule_row, rule_group, replace_index, replace_pos, group1_count, group2_count As Integer
    Dim code_full, code_family, code_replace, worksheet_write, worksheet_read, code_keep_left, code_keep_right, rule_target As String
    
    Dim index_pos(0 To 11) As Integer
    
    worksheet_read = "Transform"
    worksheet_write = "Family_List"
      
    col = 1
    row = 2
    
    code_family = "START"
    code = "START"
    
    While code <> ""
        code_family = Left(Worksheets(worksheet_read).Cells(row, col).Value, 3)
        code_full = Worksheets(worksheet_read).Cells(row, col).Value
        
        If InStr(Worksheets(worksheet_read).Cells(row, col + 1), "2700K") > 0 Then
            rule_target = "2"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), "3000K") > 0 Then
            rule_target = "3"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), "4000K") > 0 Then
            rule_target = "4"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), " HW") > 0 Then
            rule_target = "0"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), " WW") > 0 Then
            rule_target = "1"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), " NW") > 0 Then
            rule_target = "2"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row, col + 1), " CW") > 0 Then
            rule_target = "3"
            group2_count = group2_count + 1
        Else
            rule_target = ""
        End If
        
        If group1_count > 0 Or group2_count > 0 Then
            For i = 1 To Len(code_full)
    
                If Mid(code_full, i, 1) = rule_target Then
                    index_pos(i - 1) = index_pos(i - 1) + 1
                End If
            Next i
            
'            For i = LBound(index_pos) To UBound(index_pos)
'                Debug.Print i & ": " & index_pos(i)
'            Next i
        End If
        
        If code_family <> Left(Worksheets(worksheet_read).Cells(row + 1, col).Value, 3) Then
            For i = 0 To UBound(index_pos)
                If index_pos(i) = Application.WorksheetFunction.Max(index_pos) Then
                    replace_pos = i
                End If
            Next i
            Worksheets(worksheet_write).Cells(Search_Rule(code_family), 5).Value = replace_pos + 1
            If group1_count > 0 Then
                rule_group = 1
            ElseIf group2_count > 0 Then
                rule_group = 2
            ElseIf group1_count = 0 And group2_count = 0 Then
                rule_group = 0
                Worksheets(worksheet_write).Cells(Search_Rule(code_family), 5).Value = ""
            End If
            Worksheets(worksheet_write).Cells(Search_Rule(code_family), 4).Value = rule_group
            
            For i = LBound(index_pos) To UBound(index_pos)
                Debug.Print i & ": " & index_pos(i)
            Next i
            
            Erase index_pos
            group1_count = 0
            group2_count = 0
        End If

        row = row + 1
    Wend

End Sub
Sub Sort_Rules()

    Dim col_now, row_now, col_current_fam_start, row_current_fam_start, replacement_index, group1_count, group2_count, col_write, row_write, check_fail As Integer
    Dim code_full, code_family, code_replace, worksheet_write, worksheet_read, code_keep_left, code_keep_right, rule_target, rule_group As String
    
    Dim index_poll(1 To 12) As Integer
    
    worksheet_read = "New_Family_Products"
    worksheet_write = "New_family"
      
    col_now = 2
    row_now = 2
    
    col_current_fam_start = col_now
    row_current_fam_start = row_now
    
    code_family = "START"
    code_full = "START"
    
    'column index of Rule Group column in the Family_List sheet
    col_write = 4
    
    While code_full <> ""
        
'        If row_now = 138023 Then
'            row_now = row_now
'        End If
                
        
        
        code_full = Worksheets(worksheet_read).Cells(row_now, col_now).Value
        code_family = Left(code_full, 3)
        
        If InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "2700K") > 0 Then
            rule_target = "2"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "3000K") > 0 Then
            rule_target = "3"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "4000K") > 0 Then
            rule_target = "4"
            group1_count = group1_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "HW") > 0 Then
            rule_target = "0"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "WW") > 0 Then
            rule_target = "1"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "NW") > 0 Then
            rule_target = "2"
            group2_count = group2_count + 1
        ElseIf InStr(Worksheets(worksheet_read).Cells(row_now, col_now + 1), "CW") > 0 Then
            rule_target = "3"
            group2_count = group2_count + 1
        Else
            rule_target = 0
        End If
        
        If rule_target > 0 Then
            For i = 1 To 12
                If Mid(code_full, i, 1) = rule_target Then
                    index_poll(i) = index_poll(i) + 1
                End If
            Next i
        End If
        
        'For validation and monitoring purpose
        For i = LBound(index_poll) To UBound(index_poll)
            Debug.Print i & ": " & index_poll(i)
        Next i
        
        'Check if next line item belongs in the same family, if so, complete family analysis
        If code_family <> Left(Worksheets(worksheet_read).Cells(row_now + 1, col_now).Value, 3) Then
            
            
            
            'Find the most polled index position
            For i = LBound(index_poll) To UBound(index_poll)
                If index_poll(i) = Application.WorksheetFunction.Max(index_poll) Then
                    replacement_index = i
                End If
            Next i
            
            'Find which group this product family falls under
            If group1_count > group2_count Then
                rule_group = "1"
            ElseIf group2_count > group1_count Then
                rule_group = "2"
            Else
                rule_group = "NA"
            End If
                

                
            'Write result to rule form
            row_write = Search_Rule(code_family)
            
            If rule_group <> "NA" Then
                Worksheets(worksheet_write).Cells(row_write, col_write).Value = rule_group
                Worksheets(worksheet_write).Cells(row_write, col_write + 1).Value = replacement_index
            Else
                Worksheets(worksheet_write).Cells(row_write, col_write).Value = rule_group
                Worksheets(worksheet_write).Cells(row_write, col_write + 1).Value = "NA"
            End If
            
            For i = LBound(index_poll) To UBound(index_poll)
                Debug.Print i & ": " & index_poll(i)
            Next i
            
            'reiterate through last full family section and check for replacement_index validity
            While row_current_fam_start <= row_now
                If InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "2700K") > 0 Then
                    rule_target = "2"
                    'group1_count = group1_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "3000K") > 0 Then
                    rule_target = "3"
                    'group1_count = group1_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "4000K") > 0 Then
                    rule_target = "4"
                    'group1_count = group1_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "HW") > 0 Then
                    rule_target = "0"
                    'group2_count = group2_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "WW") > 0 Then
                    rule_target = "1"
                    'group2_count = group2_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "NW") > 0 Then
                    rule_target = "2"
                    'group2_count = group2_count + 1
                ElseIf InStr(Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start + 1), "CW") > 0 Then
                    rule_target = "3"
                    'group2_count = group2_count + 1
                Else
                    rule_target = 0
                End If
                
                code_full = Worksheets(worksheet_read).Cells(row_current_fam_start, col_current_fam_start).Value
                code_family = Left(code_full, 3)
                
                If Mid(code_full, replacement_index, 1) <> rule_target And rule_target > 0 Then
                    check_fail = check_fail + 1
                End If
                
                Worksheets(worksheet_write).Cells(row_write, col_write + 3).Value = check_fail
                row_current_fam_start = row_current_fam_start + 1

            Wend
            
            'Clean up variables
            Erase index_poll
            group1_count = 0
            group2_count = 0
            rule_group = ""
            check_fail = 0
            
            row_current_fam_start = row_now + 1
     

     
        End If
       
        
        
        
        row_now = row_now + 1
    Wend
        
        

  
'        If group1_count > 0 Or group2_count > 0 Then
'            For i = 1 To Len(code_full)
'
'                If Mid(code_full, i, 1) = rule_target Then
'                    index_pos(i - 1) = index_pos(i - 1) + 1
'                End If
'            Next i
'
'            For i = LBound(index_pos) To UBound(index_pos)
'                Debug.Print i & ": " & index_pos(i)
'            Next i
'        End If
'
'        If code_family <> Left(Worksheets(worksheet_read).Cells(row + 1, col).Value, 3) Then
'            For i = 0 To UBound(index_pos)
'                If index_pos(i) = Application.WorksheetFunction.Max(index_pos) Then
'                    replace_pos = i
'                End If
'            Next i
'            Worksheets(worksheet_write).Cells(Search_Rule(code_family), 5).Value = replace_pos + 1
'            If group1_count > 0 Then
'                rule_group = 1
'            ElseIf group2_count > 0 Then
'                rule_group = 2
'            ElseIf group1_count = 0 And group2_count = 0 Then
'                rule_group = 0
'                Worksheets(worksheet_write).Cells(Search_Rule(code_family), 5).Value = ""
'            End If
'            Worksheets(worksheet_write).Cells(Search_Rule(code_family), 4).Value = rule_group
'            Erase index_pos
'            group1_count = 0
'            group2_count = 0
'        End If
'
'
'    Wend

End Sub
Sub find_new_family()

    Dim family_code, code_example, worksheet_read, worksheet_write, worksheet_compare, code_compare As String
    Dim col_read, row_read, col_write, row_write, col_compare, row_compare As Integer
    Dim code_exist As Boolean
    
    worksheet_read = "Foglio1"
    worksheet_write = "New_family"
    worksheet_compare = "Family_List"
    col_read = 2
    row_read = 2
    col_write = 1
    row_write = 2
    col_compare = 1
    row_compare = 2
    code_exist = False
    

    family_code = "Start"
    code_compare = "Start"
    
    While family_code <> ""
        family_code = Left(Worksheets(worksheet_read).Cells(row_read, col_read), 3)
        If family_code <> Worksheets(worksheet_write).Cells(row_write - 1, col_write).Value Then
            
            While code_compare <> ""
                code_compare = Left(Worksheets(worksheet_compare).Cells(row_compare, col_compare).Value, 3)
                If code_compare = family_code Then
                    code_exist = True
                End If
                row_compare = row_compare + 1
            Wend
            
            If code_exist = False Then
                Worksheets(worksheet_write).Cells(row_write, col_write).Value = family_code
                Worksheets(worksheet_write).Cells(row_write, col_write + 1).Value = Worksheets(worksheet_read).Cells(row_read, col_read).Value
                Worksheets(worksheet_write).Cells(row_write, col_write + 2).Value = Worksheets(worksheet_read).Cells(row_read, col_read + 1).Value
                row_write = row_write + 1
            End If
            
            row_compare = 2
            code_exist = False
            code_compare = "Start"
            
        End If
        row_read = row_read + 1
    Wend


End Sub
Sub move_new_family_products()

    Application.ScreenUpdating = False

    Dim family_code, code_example, worksheet_read, worksheet_write, worksheet_compare, code_compare As String
    Dim col_read, row_read, col_write, row_write, col_compare, row_compare As Integer
    Dim code_exist As Boolean
    
    worksheet_read = "Foglio1"
    worksheet_write = "New_Family_Products"
    worksheet_compare = "Family_List"
    col_read = 2
    row_read = 2
    col_write = 1
    row_write = 2
    col_compare = 1
    row_compare = 2
    code_exist = False
    

    family_code = "Start"
    code_compare = "Start"
    
    While family_code <> ""
        family_code = Left(Worksheets(worksheet_read).Cells(row_read, col_read), 3)
        While code_compare <> ""
            code_compare = Left(Worksheets(worksheet_compare).Cells(row_compare, col_compare).Value, 3)
            If code_compare = family_code Then
                code_exist = True
            End If
            row_compare = row_compare + 1
        Wend
    
        If code_exist = False Then
            Worksheets(worksheet_read).Rows(row_read).Cut
            Worksheets(worksheet_write).Rows(row_write).Insert Shift:=xlDown
            If Worksheets(worksheet_read).Cells(row_read, col_read).Value = "" Then
                Worksheets(worksheet_read).Rows(row_read).Delete Shift:=xlUp
            End If
            row_write = row_write + 1
        Else
            row_read = row_read + 1
        End If
        
        row_compare = 2
        code_exist = False
        code_compare = "Start"
        
'        If family_code <> Worksheets(worksheet_write).Cells(row_write - 1, col_write).Value Then
'
'            While code_compare <> ""
'                code_compare = Left(Worksheets(worksheet_compare).Cells(row_compare, col_compare).Value, 3)
'                If code_compare = family_code Then
'                    code_exist = True
'                End If
'                row_compare = row_compare + 1
'            Wend
'
'            If code_exist = False Then
'                Worksheets(worksheet_write).Cells(row_write, col_write).Value = family_code
'                Worksheets(worksheet_write).Cells(row_write, col_write + 1).Value = Worksheets(worksheet_read).Cells(row_read, col_read).Value
'                Worksheets(worksheet_write).Cells(row_write, col_write + 2).Value = Worksheets(worksheet_read).Cells(row_read, col_read + 1).Value
'                row_write = row_write + 1
'            End If
'
'            row_compare = 2
'            code_exist = False
'            code_compare = "Start"
'
'        End If
'        row_read = row_read + 1
    Wend

    Application.ScreenUpdating = True


End Sub


Sub test()

    Dim code_keep_left, code_keep_right, code_replace, code_all As String
    Dim replace_index As Integer
    
        
    code_all = Worksheets("Test").Cells(3, 2)
    replace_index = 8
    
    code_keep_left = Left(code_all, replace_index - 1)
    code_keep_right = Mid(code_all, replace_index + 1, Len(code_all) - replace_index)
    code_replace = Mid(code_all, replace_index, 1)
    
    code_replace = code_replace & "BB"
    
    Worksheets("Test").Cells(4, 2).Value = code_keep_left & code_replace & code_keep_right
    

End Sub
Sub Find_Rule()

    Dim name As String
    Dim cell As Range
    
    
    Set cell = Range("A3:A500").FIND("AB")
    Debug.Print cell.Address

End Sub

Sub clone3500k()

    Application.ScreenUpdating = False

    Dim col, row, row_write As Integer
'    Dim description, check_blank As String
    
    worksheet_read = "PriceList_v1"
    worksheet_write = "PriceList_v1_3500K"
      
    col = 3
    row = 2
    
    row_write = 2
    
'    description = "START"
    
    While Worksheets(worksheet_read).Cells(row, col).Value <> ""
'        description = Worksheets(worksheet_read).Cells(row, col).Value
        
        If InStr(Worksheets(worksheet_read).Cells(row, col).Value, "3000K") > 0 Then
            Worksheets(worksheet_read).Rows(row).Copy
            Worksheets(worksheet_write).Rows(row_write).Insert Shift:=xlDown
            row_write = row_write + 1
        End If

        row = row + 1
    Wend

    Application.ScreenUpdating = True


End Sub
