Sub Step1ClearContents()
Call clearcontents
Call DatavalidationPersonal
Call DatavalidationCUSTTYPE
End Sub


Sub clearcontents()

With ThisWorkbook.Worksheets("CASAsheet1").UsedRange
    'With .Cells(1, 1).CurrentRegion
        'With .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)
            .Cells.clearcontents
            .Cells.Interior.Pattern = xlNone
            .Cells.Font.ColorIndex = xlAutomatic
            .Cells.Validation.Delete
            .Columns("I").NumberFormat = "General"
            .Columns("J").NumberFormat = "General"
            .Cells(1, 1) = "PARTY_NUMBER"
            .Cells(1, 2) = "ID Indicator(Personal = ID, CI, PP, ZZ, OTW,Company= BR, OT, XX)"
            .Cells(1, 3) = "PARTY_NUMBER(First 10 to unlimited digits)"
            .Cells(1, 4) = "Suffix(Left 3 digits for BR company)"
            .Cells(1, 5) = "Personal?"
            .Cells(1, 6) = "monthly_average"
            .Cells(1, 7) = "FULL_NAME"
            .Cells(1, 8) = "OCCUPATION"
            .Cells(1, 9) = "OFFICER_CD"
            .Cells(1, 10) = "OFFICER_SUB_CD"
            .Cells(1, 11) = "RH"
            .Cells(1, 12) = "Customer Type Original"
            .Cells(1, 13) = "Customer Type 1st update"
            .Cells(1, 14) = "Customer Type Top20 Only"
            .Cells(1, 15) = "Customer Type Final"
            .Cells(1, 16) = "% of Total CASA"
            .Cells(1, 17) = "% of Total CASA (Without Personal)"
            .Cells(1, 18) = "Bills - Amount Range"
            .Cells(1, 19) = "Non-Bills Borrowing - Amount Range"
            .Cells(1, 20) = "Non-Borrowing - Amount Range"
            .Cells(1, 21) = "Tier Interest Client"
            .Cells(1, 22) = "DS Direct GROUP ID"
            .Cells(1, 23) = "% of Total CASA of Each Zone"
            .Cells(19, 26) = "Step5 please go to top 20 file"
        'End With
    'End With
End With


'For Each ws In Worksheets
    'If ws.Name = "TierInt" Then
        'Application.DisplayAlerts = False
        'Sheets("TierInt").Delete
        'Application.DisplayAlerts = True
    'End If
'Next


'Dim ws As Worksheets

Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "CASASheet1" Or ws.Name = "DataValidation" Then
    
    Else: ws.Delete
    End If
Next
Application.DisplayAlerts = True


End Sub


Sub Step2ImportRaw()

Call CASAImportRawFile
Call addZerotoBlankAmountinRaw
Call breakpartyNo
Call SumAvgAmountInRaw
Call removeRAWDuplicate
Call CASAvlookup

End Sub

Sub Step10Calculation()

Call PercentageOfCASA
Call PercentageOfCASAWithourPersonal
Call AmountRange
Call TierIntCSV_Import
Call TierIntvlookupInCASA
Call AddTrimCIF
Call ChangeDatavalidationCUSTTYPE

End Sub


Sub addZerotoBlankAmountinRaw()
   'Dim PasteStart As Range
   Dim a As Integer
    
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    For a = 2 To NumRows
    
    Range("F" & a).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
    
    If (Range("F" & a).Value) = "" Then
    Range("F" & a).Value = "0"
    
    End If
            
    Next
    
End Sub

Sub Step4FirstChangeCUSTTYPEold()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer

    Counter = 0
    NumRows = Range("C1", Range("C1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    
    'if (its br and its multiple br no)or (its OT and its multiple OT)
    If (Range("B" & a).Value = "BR" And Range("C" & a).Value = Range("C" & a + 1).Value) Or (Range("B" & a).Value = "OT" And Range("C" & a).Value = Range("C" & a + 1).Value) Then
    'If Range("C" & a).Value = Range("C" & a + 1).Value Then
    
        Counter = 0
        StartduplicateNum = 0
       'Range("A" & a).Offset(0, 7).value = "kaka"
        StartduplicateNum = a
        
        'Find the total no of parent and son row (same br no)
        'Do Until Left(Range("A" & a).Value, 10) <> Left(Range("A" & a + 1).Value, 10)
        Do Until Range("C" & a).Value <> Range("C" & a + 1).Value
        Counter = Counter + 1
        a = a + 1
        'MsgBox (Left(Range("A" & a).Value, 10))
        'MsgBox ("StartduplicateNum =" & StartduplicateNum)
        'MsgBox ("counter=" & Counter)
        'MsgBox ("a=" & a)
        Loop
    
               'MsgBox ("working on parent")
               If Cells(StartduplicateNum, "M") = "Bills Customer" Or Cells(StartduplicateNum, "M") = "Non-Bills Borrowing Customer" Then  'do nothing"
               Else:
                    For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
                        If Cell.Value = "Bills Customer" Or Cell.Value = "Non-Bills Borrowing Customer" Then
                        'parent = the value of son
                        Cells(StartduplicateNum, "M").Value = Cell.Value
                        Cells(StartduplicateNum, "M").Interior.ColorIndex = 6 'yellow
                        
                        End If
                    Next
               End If
               
               'MsgBox ("working on son")
               For Each Cell In Range("M" & StartduplicateNum + 1 & ":" & "M" & StartduplicateNum + Counter)
                    If Cell.Value = "Bills Customer" Or Cell.Value = "Non-Bills Borrowing Customer" Then  'do nothing
                    ' son = the value of parent
                    Else:
                    Cell.Value = Cells(StartduplicateNum, "M").Value
                    Cell.Interior.ColorIndex = 6 'yellow
                    End If
              Next
      
    Else 'do nothing
    
    'MsgBox "The loop made " & counter & " repetitions."
    End If
    'MsgBox (Left(Range("A" & a).Value, 10))
    'MsgBox ("a1=" & a)
    Next
End Sub

Sub Step4FirstChangeCUSTTYPEnew1()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer
Dim iRow As Long
Dim bln As String
Dim bln1 As String
Dim var As Variant
Dim NumRows As Integer
Dim i As Integer

    Counter = 0
    NumRows = Range("C1", Range("C1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    'MsgBox (NumRows)
    
    'if (its br and its multiple br no)or (its OT and its multiple OT)
    If (Range("B" & a).Value = "BR" And Range("C" & a).Value = Range("C" & a + 1).Value) Or (Range("B" & a).Value = "OT" And Range("C" & a).Value = Range("C" & a + 1).Value) Then
    'If Range("C" & a).Value = Range("C" & a + 1).Value Then
        bln = ""
        bln1 = ""
        Counter = 0
        StartduplicateNum = 0
       'Range("A" & a).Offset(0, 7).value = "kaka"
        StartduplicateNum = a
        
        'Find the total no of parent and son row (same br no)
        'Do Until Left(Range("A" & a).Value, 10) <> Left(Range("A" & a + 1).Value, 10)
        Do Until Range("C" & a).Value <> Range("C" & a + 1).Value
        Counter = Counter + 1
        a = a + 1
        'MsgBox (Left(Range("A" & a).Value, 10))
        'MsgBox ("StartduplicateNum =" & StartduplicateNum)
        'MsgBox ("counter=" & Counter)
        'MsgBox ("a=" & a)
        Loop
    
    
    'For iRow = StartduplicateNum To StartduplicateNum + Counter
          'For every cell that is not empty, search through the first column in each worksheet in the
          'workbook for a value that matches that cell value.

          Set m_rnCheck = Range("L" & StartduplicateNum & ":" & "L" & StartduplicateNum + Counter)
          'If Not IsEmpty(Cells(iRow, "M")) Then
            'For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             'For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                'bln = False
                
            With m_rnCheck
                Set m_rnFind = .Find(What:="Bills Customer", LookIn:=xlFormulas)
                
                If Not m_rnFind Is Nothing Then
                'm_stAddress = m_rnFind.Address
                bln = "Bills Customer"
                
                End If
                
                Set m_rnFind1 = .Find(What:="Non-Bills Borrowing Customer", LookIn:=xlFormulas)
                
                If Not m_rnFind1 Is Nothing Then
                'm_stAddress = m_rnFind.Address
                bln1 = "Non-Bills Borrowing Customer"
               
                End If
                
                'Else
                
                    'Set m_rnFind = .Find(What:="Non-Bills Borrowing Customer", LookIn:=xlFormulas)
                
                    'If Not m_rnFind Is Nothing Then
                    ''m_stAddress = m_rnFind.Address
                    'bln = "Non-Bills Borrowing Customer"
                
                            
                    'Else
                
                    'bln = ""
                
                    'End If
                
                'End If

                'If you find a matching value, indicate success by setting bln to true and exit the loop;
                'otherwise, continue searching until you reach the end of the workbook.
             
                                       
             End With
             'Next
             'Next iSheet
          'End If
          
          'If you do not find a matching value, do not bold the value in the original list;
          'if you do find a value, bold it.
          'MsgBox ("BB")
          
          For Each Cell In Range("C" & StartduplicateNum & ":" & "C" & StartduplicateNum + Counter)
            
             Cell.Interior.ColorIndex = 6 'yellow
          
          Next
                  
                    
          If (bln <> "" And bln1 <> "") Or (bln <> "" And bln1 = "") Then
             For Each Cell In Range("L" & StartduplicateNum & ":" & "L" & StartduplicateNum + Counter)
             If Cell.Value = "Non-Borrowing Customers" Then
             Cell.Offset(0, 1).Value = bln
             Cell.Offset(0, 1).Interior.ColorIndex = 45 'orange
             'MsgBox (StartduplicateNum & "colour")
             Else
             
             Cell.Offset(0, 1).Value = Cell.Value
             End If
             Next
           
           ElseIf bln = "" And bln1 <> "" Then
             For Each Cell In Range("L" & StartduplicateNum & ":" & "L" & StartduplicateNum + Counter)
             If Cell.Value = "Non-Borrowing Customers" Then
             Cell.Offset(0, 1).Value = bln1
             Cell.Offset(0, 1).Interior.ColorIndex = 45 'orange
             'MsgBox (StartduplicateNum & "colour")
             Else
             Cell.Offset(0, 1).Value = Cell.Value
             End If
             Next
             
          'bln= "" and bln1 = ""
           Else
             For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             Cell.Value = Range("L" & StartduplicateNum).Value
             'cell.Interior.ColorIndex = 45 'orange
             Next
           End If
            
       'Next iRow

      
    Else 'do nothing
    'MsgBox ("is not duplicate party number")
    'MsgBox "The loop made " & counter & " repetitions."
    Range("M" & a).Value = Range("L" & a).Value
    End If
   
    'MsgBox ("a1=" & a)
    Next
End Sub

Sub Step4FirstChangeCUSTTYPEnew()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer
Dim iRow As Long
Dim bln As String
Dim var As Variant
Dim NumRows As Integer
Dim i As Integer

    Counter = 0
    NumRows = Range("C1", Range("C1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    'MsgBox (NumRows)
    
    'if (its br and its multiple br no)or (its OT and its multiple OT)
    If (Range("B" & a).Value = "BR" And Range("C" & a).Value = Range("C" & a + 1).Value) Or (Range("B" & a).Value = "OT" And Range("C" & a).Value = Range("C" & a + 1).Value) Then
    'If Range("C" & a).Value = Range("C" & a + 1).Value Then
        bln = ""
        Counter = 0
        StartduplicateNum = 0
       'Range("A" & a).Offset(0, 7).value = "kaka"
        StartduplicateNum = a
        
        'Find the total no of parent and son row (same br no)
        'Do Until Left(Range("A" & a).Value, 10) <> Left(Range("A" & a + 1).Value, 10)
        Do Until Range("C" & a).Value <> Range("C" & a + 1).Value
        Counter = Counter + 1
        a = a + 1
        'MsgBox (Left(Range("A" & a).Value, 10))
        'MsgBox ("StartduplicateNum =" & StartduplicateNum)
        'MsgBox ("counter=" & Counter)
        'MsgBox ("a=" & a)
        Loop
    
    
    'For iRow = StartduplicateNum To StartduplicateNum + Counter
          'For every cell that is not empty, search through the first column in each worksheet in the
          'workbook for a value that matches that cell value.

          Set m_rnCheck = Range("L" & StartduplicateNum & ":" & "L" & StartduplicateNum + Counter)
          'If Not IsEmpty(Cells(iRow, "M")) Then
            'For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             'For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                'bln = False
                
            With m_rnCheck
                Set m_rnFind = .Find(What:="Bills Customer", LookIn:=xlFormulas)
                
                If Not m_rnFind Is Nothing Then
                'm_stAddress = m_rnFind.Address
                bln = "Bills Customer"
                
               
                Else
                
                    Set m_rnFind = .Find(What:="Non-Bills Borrowing Customer", LookIn:=xlFormulas)
                
                    If Not m_rnFind Is Nothing Then
                    'm_stAddress = m_rnFind.Address
                    bln = "Non-Bills Borrowing Customer"
                
                            
                    Else
                
                    bln = ""
                
                    End If
                
                End If

                'If you find a matching value, indicate success by setting bln to true and exit the loop;
                'otherwise, continue searching until you reach the end of the workbook.
             
                         
             
             
             End With
             'Next
             'Next iSheet
          'End If
          
          'If you do not find a matching value, do not bold the value in the original list;
          'if you do find a value, bold it.
          'MsgBox ("BB")
          If bln <> "" Then
             For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             Cell.Value = bln
             Cell.Interior.ColorIndex = 39 'purple
             'MsgBox (StartduplicateNum & "colour")
             Next
           Else
             
             For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             Cell.Value = Range("L" & StartduplicateNum).Value
             Cell.Interior.ColorIndex = 39 'purple
             Next
           End If
            
       'Next iRow

      
    Else 'do nothing
    'MsgBox ("is not duplicate party number")
    'MsgBox "The loop made " & counter & " repetitions."
    Range("M" & a).Value = Range("L" & a).Value
    End If
   
    'MsgBox ("a1=" & a)
    Next
End Sub


Sub CASAImportRawFile()

Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Sheet As Worksheet
Dim PasteStart As Range

Set wb1 = ActiveWorkbook
Set PasteStart = [CASASheet1!A1]

FileToOpen = Application.GetOpenFilename _
(Title:="******Please choose THIS MONTH CASA_raw_YYYYMM*****", _
FileFilter:="Report Files *.csv(*.csv),")

If FileToOpen = False Then
    MsgBox "No File Specified.", vbExclamation, "ERROR"
    Exit Sub
Else
    Set wb2 = Workbooks.Open(filename:=FileToOpen)

    For Each Sheet In wb2.Sheets
        With Range("A1", Range("A" & Rows.Count).End(xlUp))
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            .Offset(0, 0).EntireColumn.Copy PasteStart.Range("A1")
            .Offset(0, 2).EntireColumn.Copy PasteStart.Range("F1")
            
                           
        End With
    Next Sheet
End If

    wb2.Close
    
End Sub
    
    
  Sub breakpartyNo()
   Dim PasteStart As Range
   Dim a As Integer
    
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    For a = 2 To NumRows
    
    If Left(Range("A" & a).Value, 2) = "BR" Then
    
    'With Worksheets("CASASheet1").Range("A1", "A" & NumRows)
    With Worksheets("CASASheet1").Range("A" & a)
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            .Offset(0, 1) = Left(Range("A" & a).Value, 2)
            .Offset(0, 2) = Left(Range("A" & a).Value, 10)
            .Offset(0, 3).NumberFormat = "@"
            .Offset(0, 3) = Right(Range("A" & a).Value, 3)
               
    End With
    ElseIf Left(Range("A" & a).Value, 3) = "OTW" Then
    With Worksheets("CASASheet1").Range("A" & a)
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            .Offset(0, 1) = Left(Range("A" & a).Value, 3)
            .Offset(0, 2) = Range("A" & a).Value
                 
    End With
    
    Else
    
    With Worksheets("CASASheet1").Range("A" & a)
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            .Offset(0, 1) = Left(Range("A" & a).Value, 2)
            .Offset(0, 2) = Range("A" & a).Value
            
               
    End With
    End If
    
    
    
    If Range("B" & a).Value = "ID" Or Range("B" & a).Value = "CI" Or Range("B" & a).Value = "PP" Or Range("B" & a).Value = "ZZ" Or Range("B" & a).Value = "OTW" Then
    With Worksheets("CASASheet1").Range("A" & a)
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            .Offset(0, 4) = "Y"
                        
    End With
    
    ElseIf Range("B" & a).Value = "BR" Or Range("B" & a).Value = "OT" Or Range("B" & a).Value = "XX" Then
    With Worksheets("CASASheet1").Range("A" & a)
            .Offset(0, 4) = "N"
            
              
            
    End With
     Else
     
    With Worksheets("CASASheet1").Range("A" & a)
            .Offset(0, 4) = CVErr(xlErrNA)
            
              
            
    End With
     
     
    End If
    
    Next
    
End Sub

Sub SumAvgAmountInRaw()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer

    Counter = 0
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    For a = 1 To NumRows
    
    'if (its br and its multiple br no)or (its <>BR and its multiple rows)
    'If (Left(Range("A" & a).Value, 2) = "BR" And Left(Range("A" & a).Value, 10) = Left(Range("A" & a + 1).Value, 10)) Or (Left(Range("A" & a).Value, 2) <> "BR" And Range("A" & a).Value = Range("A" & a + 1).Value) Then
    If Range("A" & a).Value = Range("A" & a + 1).Value Then
    
        Counter = 0
        StartduplicateNum = 0
       'Range("A" & a).Offset(0, 7).value = "kaka"
        StartduplicateNum = a
        
        'Find the total no of parent and son row (same party no)
        Do Until Range("A" & a).Value <> Range("A" & a + 1).Value
        Counter = Counter + 1
        a = a + 1
        'MsgBox (Left(Range("A" & a).Value, 10))
        'MsgBox ("StartduplicateNum =" & StartduplicateNum)
        'MsgBox ("counter=" & Counter)
        'MsgBox ("a=" & a)
        Loop
    
               'MsgBox ("working on parent")
               'If Cells(StartduplicateNum, 8) = "Bills Customer" Or Cells(StartduplicateNum, 8) = "Non-Bills Borrowing Customer" Then  'do nothing"
               'Else:
                    Sum = 0
                    For Each Cell In Range("F" & StartduplicateNum & ":" & "F" & StartduplicateNum + Counter)
                        Sum = Sum + Cell.Value
                        'parent = the value of son
                        'Cells(StartduplicateNum, 8).Value = cell.Value
                        'Cells(StartduplicateNum, "F").Interior.ColorIndex = 42 'blue
                    Next
                    
                    For Each Cell In Range("F" & StartduplicateNum & ":" & "F" & StartduplicateNum + Counter)
                    Cell.Value = Sum
                    Cell.Interior.ColorIndex = 42 'blue
                    Next
                    
               'End If
               
               'MsgBox ("working on son")
               'For Each cell In Range("H" & StartduplicateNum + 1 & ":" & "H" & StartduplicateNum + Counter)
                    'If cell.Value = "Bills Customer" Or cell.Value = "Non-Bills Borrowing Customer" Then  'do nothing
                    ' son = the value of parent
                    'Else:
                    'cell.Value = Cells(StartduplicateNum, 8).Value
                    'cell.Interior.ColorIndex = 6 'yellow
                    'End If
              'Next
      
    Else 'do nothing
    
    'MsgBox "The loop made " & counter & " repetitions."
    End If
    'MsgBox (Left(Range("A" & a).Value, 10))
    'MsgBox ("a1=" & a)
    Next
End Sub

Sub removeRAWDuplicate()
 'removeDuplicate Macro
 'Columns("A:A").Select
 Worksheets("CASASheet1").Range("A1", Range("F" & Rows.Count).End(xlUp)).RemoveDuplicates Columns:=Array(1), _
 Header:=xlNo
 Range("A1").Select
 End Sub

Sub Step3CASAvlookupPreviousCUSTTYPE()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose PREVIOUS MONTH cust_type_YYYYMM*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in cust_type_YYYYMM
Set mySheetName = myFileName.ActiveSheet

'find last row in Column A ("PARTY_NUMBER")in cust_type_YYYYMM
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in cust_type_YYYYMM
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        
        If IsError(currentSheet.Cells(i, "L").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            .Cells(i, "L") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            '.Cells(i, "L") = CVErr(xlErrNA)
            '.Cells(i, "M") = CVErr(xlErrNA)
            '.Cells(i, "L").Interior.ColorIndex = 8
        End If
        
        End If
    End With
Next i

myFileName.Close


End Sub


Sub CASAvlookup()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose THIS MONTH Cust_info_YYYYMM*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Cust_Info_YYYYMM
Set mySheetName = myFileName.ActiveSheet

'find last row in Column A ("PARTY_NUMBER")in Cust_Info_YYYYMM
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in Cust_Info_YYYYMM
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            .Cells(i, "G") = mySheetName.Cells(matchRow, "B").Value
            .Cells(i, "H") = mySheetName.Cells(matchRow, "D").Value
            .Cells(i, "I") = mySheetName.Cells(matchRow, "E").Value
            .Cells(i, "J") = mySheetName.Cells(matchRow, "F").Value
            .Cells(i, "K") = mySheetName.Cells(matchRow, "G").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "G") = CVErr(xlErrNA)
            .Cells(i, "H") = CVErr(xlErrNA)
            .Cells(i, "I") = CVErr(xlErrNA)
            .Cells(i, "J") = CVErr(xlErrNA)
            .Cells(i, "K") = CVErr(xlErrNA)
        End If

    End With
Next i


myFileName.Close

'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose THIS MONTH cust_type_YYYYMM*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in cust_type_YYYYMM
Set mySheetName = myFileName.ActiveSheet

'find last row in Column A ("PARTY_NUMBER")in cust_type_YYYYMM
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in cust_type_YYYYMM
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            .Cells(i, "L") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "L") = CVErr(xlErrNA)
            '.Cells(i, "M") = CVErr(xlErrNA)
        End If

    End With
Next i

myFileName.Close


End Sub

Sub Step9CASATop20changeCUSTTYPEnew1()


Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet

Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls), *.xls", Title:="******Please choose your THIS MONTH CASA TOP 20 Workbook*****")
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls;*.xlsx), *.xls; *.xlsx", Title:="******Please choose your THIS MONTH(after add co) CASA TOP 20 Workbook*****")
filename = Application.GetOpenFilename(Title:="******Please choose your THIS MONTH(after add co) CASA TOP 20 Workbook*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)


'set searched worksheet in your CASA TOP 20 Workbook
Set mySheetName = myFileName.ActiveSheet

'find last row in Column A ("PARTY_NUMBER")in your CASA TOP 20 Workbook
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in your CASA TOP 20 Workbook
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)) Then
            matchRow = Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)
            .Cells(i, "N") = mySheetName.Cells(matchRow, "M").Value
            '.Cells(i, "N").Interior.ColorIndex = 22 'pink
            '.Cells(i, "A").Interior.ColorIndex = 22 'pink
            
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "N") = CVErr(xlErrNA)
            
        End If
        
        currentSheet.Cells(i, "O") = .Cells(i, "M")
        If IsError(currentSheet.Cells(i, "N").Value) Then
         
         Else
           
           If currentSheet.Cells(i, "M") = .Cells(i, "N") Then
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
           Else
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
                      currentSheet.Cells(i, "O").Interior.ColorIndex = 17 ' purple
                      
           End If
        End If
    End With
Next i

myFileName.Close

End Sub

Sub Step7CheckFirst20NonBorrowing()


Call CheckFirst20NonBorrowingSorting
Call CheckFirst20NonBorrowingPart1
Call CheckFirst20NonBorrowingPart2

End Sub

Sub step8clearTOP20andfinalCUSTTYPE()

With Worksheets("CASASheet1")
    .Columns(14).clearcontents
    .Columns(14).Interior.Pattern = xlNone
    .Columns(15).clearcontents
    .Columns(15).Interior.Pattern = xlNone
    .Cells(1, 14) = "Customer Type Top20 Only"
    .Cells(1, 15) = "Customer Type Final"
    
    '.Range("N2:N" & .Range("N2").End(xlDown).Row).ClearContents
    '.Range("N2:N" & .Range("N2").End(xlDown).Row).Interior.Pattern = xlNone
    '.Range("N2:O" & .Range("O2").End(xlDown).Row).ClearContents
    '.Range("N2:O" & .Range("O2").End(xlDown).Row).Interior.Pattern = xlNone
    
End With

End Sub

Sub Step9CASATop20changeCUSTTYPE()


Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet

Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls), *.xls", Title:="******Please choose your CASA TOP 20 Workbook*****")
filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls;*.xlsx), *.xls; *.xlsx", Title:="******Please choose your CASA TOP 20 Workbook*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in your CASA TOP 20 Workbook
Set mySheetName = myFileName.ActiveSheet

'find last row in Column A ("PARTY_NUMBER")in your CASA TOP 20 Workbook
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in your CASA TOP 20 Workbook
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(Trim(.Cells(i, "A")), myRangeName, 0)) Then
            matchRow = Application.Match(Trim(.Cells(i, "A")), myRangeName, 0)
            .Cells(i, "N") = mySheetName.Cells(matchRow, "J").Value
            '.Cells(i, "N").Interior.ColorIndex = 22 'pink
            .Cells(i, "A").Interior.ColorIndex = 22 'pink
            
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "N") = CVErr(xlErrNA)
            
        End If
        
        currentSheet.Cells(i, "O") = .Cells(i, "M")
        If IsError(currentSheet.Cells(i, "N").Value) Then
         
         Else
           
           If currentSheet.Cells(i, "M") = .Cells(i, "N") Then
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
           Else
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
                      currentSheet.Cells(i, "O").Interior.ColorIndex = 17 ' purple
                      
           End If
        End If
    End With
Next i

myFileName.Close

End Sub

Sub PercentageOfCASA()

Dim x As Integer
Dim a As Integer
Dim NumRows As Integer
Dim SumOfAvgAmount As Double

    SumOfAvgAmount = 0
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    SumOfAvgAmount = SumOfAvgAmount + Range("F" & a).Value
    Next
    'MsgBox (SumOfAvgAmount)
    
    
    For a = 2 To NumRows
    For Each Cell In Range("P" & a & ":" & "P" & a)
                    Cell.Value = Cell.Offset(0, -10) / SumOfAvgAmount
                    Cell.NumberFormat = "0.00%"
    Next
    Next
    
End Sub

Sub PercentageOfCASAWithourPersonal()

Dim x As Integer
Dim a As Integer
Dim NumRows As Integer
Dim SumOfAvgAmountwithoutpersonal As Double

    SumOfAvgAmount = 0
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    
    If Range("E" & a) = "N" Then
    SumOfAvgAmountwithoutpersonal = SumOfAvgAmountwithoutpersonal + Range("F" & a).Value
    Else 'nth
    End If
    Next
    'MsgBox (SumOfAvgAmount)
    
    
    For a = 2 To NumRows
    For Each Cell In Range("Q" & a & ":" & "Q" & a)
                    Cell.Value = Cell.Offset(0, -11) / SumOfAvgAmountwithoutpersonal
                    Cell.NumberFormat = "0.00%"
    Next
    Next
    
End Sub

Sub AmountRange()

Dim x As Integer
Dim a As Integer


    
   
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
    
    'if (its br and its multiple br no)or (its <>BR and its multiple rows)
    'If (Left(Range("A" & a).Value, 2) = "BR" And Left(Range("A" & a).Value, 10) = Left(Range("A" & a + 1).Value, 10)) Or (Left(Range("A" & a).Value, 2) <> "BR" And Range("A" & a).Value = Range("A" & a + 1).Value) Then
    
    For Each Cell In Range("O" & a & ":" & "O" & a)
    If Cell.Value = "Bills Customer" Then
                        
                        Select Case True
                            Case (Range("F" & a).Value < 10000000)
                            Cell.Offset(0, 3).Value = "<10mn"
                            Case (Range("F" & a).Value >= 10000000 And Range("F" & a).Value < 100000000)
                            Cell.Offset(0, 3).Value = "10mn-100mn"
                            Case (Range("F" & a).Value >= 100000000)
                            Cell.Offset(0, 3).Value = ">100mn"
                        End Select
                        
                    'Next
    ElseIf Cell.Value = "Non-Bills Borrowing Customer" Then
                        
                       
                        Select Case True
                            Case (Range("F" & a).Value < 10000000)
                            Cell.Offset(0, 4).Value = "<10mn"
                            Case (Range("F" & a).Value >= 10000000 And Range("F" & a).Value < 100000000)
                            Cell.Offset(0, 4).Value = "10mn-100mn"
                            Case (Range("F" & a).Value >= 100000000 And Range("F" & a).Value < 200000000)
                            Cell.Offset(0, 4).Value = "100mn-200mn"
                            Case (Range("F" & a).Value >= 200000000 And Range("F" & a).Value < 300000000)
                            Cell.Offset(0, 4).Value = "200mn-300mn"
                            Case (Range("F" & a).Value >= 300000000 And Range("F" & a).Value < 400000000)
                            Cell.Offset(0, 4).Value = "300mn-400mn"
                            Case (Range("F" & a).Value >= 400000000)
                            Cell.Offset(0, 4).Value = ">400mn"
                            
                        End Select
                    
                    
                    
    ElseIf Cell.Value = "Non-Borrowing Customers" Then
    
                        Select Case True
                            Case (Range("F" & a).Value < 10000000)
                            Cell.Offset(0, 5).Value = "<10mn"
                            Case (Range("F" & a).Value >= 10000000 And Range("F" & a).Value < 100000000)
                            Cell.Offset(0, 5).Value = "10mn-100mn"
                            Case (Range("F" & a).Value >= 100000000 And Range("F" & a).Value < 200000000)
                            Cell.Offset(0, 5).Value = "100mn-200mn"
                            Case (Range("F" & a).Value >= 200000000 And Range("F" & a).Value < 300000000)
                            Cell.Offset(0, 5).Value = "200mn-300mn"
                            Case (Range("F" & a).Value >= 300000000 And Range("F" & a).Value < 400000000)
                            Cell.Offset(0, 5).Value = "300mn-400mn"
                            Case (Range("F" & a).Value >= 400000000)
                            Cell.Offset(0, 5).Value = ">400mn"
                            
                        End Select
                        
    Else ' nothing
    
    
    End If
    'MsgBox (Left(Range("A" & a).Value, 10))
    'MsgBox ("a1=" & a)
    Next
   Next



    
End Sub

Sub TierIntCSV_Import()
Dim ws As Worksheet, strFile As String

Set ws = ActiveWorkbook.Sheets("CASASheet1") 'set to current worksheet name

strFile = Application.GetOpenFilename("Text Files (*.txt),*.txt", , "Please select **CBD12023_CA_TX_REPORT_YYYYMMDD.TXT***")

Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws1.Name = "TierInt"

With ws1.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws1.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileSemicolonDelimiter = True
     .Refresh
End With
End Sub

Sub removeTierIntDuplicate()

'removeDuplicate Macro
 
 Worksheets("TierInt").Select
 'Columns("A:A").Select
 ThisWorkbook.Worksheets("TierInt").Range("A3", Range("U" & Rows.Count).End(xlUp)).RemoveDuplicates Columns:=Array(5), Header:=xlNo
 Range("A1").Select
 End Sub
 
Sub TierIntvlookupInCASA()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long
Dim mySheetlastRow          As Long

'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose Cust_info_YYYYMM*****")

'set our workbook and open it
'Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in TierInt
Set mySheetName = ThisWorkbook.Worksheets("TierInt")

'find last row in Column E ("PARTY_NUMBER")in TierInt
mySheetlastRow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in TierInt
Set myRangeName = mySheetName.Range("E1:E" & mySheetlastRow)


For a = 3 To mySheetlastRow
    With mySheetName
        For Each Cell In Range("E" & a & ":" & "E" & a)
                    Cell.Value = Trim(Cell.Value)
                    
        Next

    End With
  Next

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)



For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            .Cells(i, "U") = mySheetName.Cells(matchRow, "E").Value
          
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "U") = CVErr(xlErrNA)
            
        End If

    End With
Next i

Worksheets("CASASheet1").Select

End Sub

Sub CheckFirst20NonBorrowingPart1OLD()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer
Dim Counter1 As Integer
Dim iRow As Long
Dim bln As String
Dim bln1 As String
Dim var As Variant
Dim NumRows As Integer
Dim i As Integer

Dim lastrow As Range
Dim m_rnCheck As Range
Dim m_rnCheck1 As Range
Dim m_rnFind As Range
Dim m_rnFind1 As Range
Dim m_stAddress As String
Dim m_stAddress1 As String
Dim lastrowNonBorrowing As Range

    
    
    Counter = 0
    NumRows = Range("C1", Range("C1").End(xlDown)).Rows.Count
    
    'For a = 2 To NumRows
    'MsgBox (NumRows)
       
           
          
          Set m_rnCheck1 = Range("E" & 1 & ":" & "E" & NumRows)
          
          With m_rnCheck1
             Set m_rnFind1 = .Find(What:="N") 'LookIn:=xlFormulas)
                    If Not m_rnFind1 Is Nothing Then
                    m_stAddress1 = m_rnFind1.Address
                    'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                    m_rnFind1.Offset(0, -4).Font.Color = vbRed
                    
                    
                    'MsgBox ("first" & m_stAddress)
                    
                    'Unhide the column, and then find the next X.
                        Do
                                                       
                            Set lastrowNonBorrowing1 = m_rnFind1
                            Set m_rnFind1 = .FindNext(m_rnFind1)
                            
                            'If m_rnFind.Offset(0, -10) = "N" Then
                            'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                            m_rnFind1.Offset(0, -4).Font.Color = vbRed
                            Counter1 = Counter1 + 1
                            'End If
                            
                             'MsgBox (m_rnFind)
                            'MsgBox ("m_rnFind.Address" & m_rnFind.Address)
                             'MsgBox ("Counter" & Counter)
                        Loop While Not m_rnFind1 Is Nothing And m_rnFind1.Address <> m_stAddress1 And Counter1 <> 19
                    End If
            End With
          
          
                  
          
          
          Set m_rnCheck = Range("O" & 1 & ":" & "O" & NumRows)
          
          'If Not IsEmpty(Cells(iRow, "M")) Then
            'For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             'For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                'bln = False
            'MsgBox (NumRows)
              
                    
           With m_rnCheck
             Set m_rnFind = .Find(What:="Non-Borrowing Customers") 'LookIn:=xlFormulas)
           
                    If Not m_rnFind Is Nothing Then
                    
                    'If m_rnFind.Offset(0, -10) = "N" Then
                      m_stAddress = m_rnFind.Address
                      'MsgBox ("haha")
                      m_rnFind.Offset(0, -14).Font.Color = vbCyan
                    'Else
                    'Do
                    'Set lastrow = m_rnFind
                    'Set m_rnFind = .FindNext(m_rnFind)
                    
                    'MsgBox (lastrow.Row)
                    'Loop While m_rnFind.Offset(0, -10) = "N"
                    'm_stAddress = m_rnFind.Address
                    'MsgBox ("hehe")
                    
                    'm_rnFind.Offset(0, -14).Font.Color = vbGreen
                    'End If
                    
                    'MsgBox ("first" & m_stAddress)
                    
                    'Unhide the column, and then find the next X.
            
                        Do
                                                       
                            Set lastrowNonBorrowing = m_rnFind
                            Set m_rnFind = .FindNext(m_rnFind)
                            
                            If m_rnFind.Offset(0, -10) = "N" Then
                            'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                            m_rnFind.Offset(0, -14).Font.Color = vbBlue
                            Counter = Counter + 1
                            End If
                            
                             'MsgBox (m_rnFind)
                            'MsgBox ("m_rnFind.Address" & m_rnFind.Address)
                             'MsgBox ("Counter" & Counter)
                        Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress And Counter <> 19
                    End If
            
            End With
            MsgBox (Counter)
            If Counter < 19 Then
            MsgBox ("There are less than 20 Non-Borrowing Customers")
            End If
            
            'MsgBox ("last row" & lastrowNonBorrowing.Row)
            
     'Next
End Sub

Sub CheckFirst20NonBorrowingPart1()

Dim x As Integer
Dim a As Integer
Dim StartduplicateNum As Integer
Dim Counter As Integer
Dim Counter1 As Integer
Dim iRow As Long
Dim bln As String
Dim bln1 As String
Dim var As Variant
Dim NumRows As Integer
Dim i As Integer

Dim lastrow As Range
Dim m_rnCheck As Range
Dim m_rnCheck1 As Range
Dim m_rnFind As Range
Dim m_rnFind1 As Range
Dim m_stAddress As String
Dim m_stAddress1 As String
Dim lastrowNonBorrowing As Range

    
    
    Counter = 0
    NumRows = Range("C1", Range("C1").End(xlDown)).Rows.Count
    
    'For a = 2 To NumRows
    'MsgBox (NumRows)
       
           
          
          Set m_rnCheck1 = Range("E" & 1 & ":" & "E" & NumRows)
          
          With m_rnCheck1
             Set m_rnFind1 = .Find(What:="N") 'LookIn:=xlFormulas)
                    If Not m_rnFind1 Is Nothing Then
                    m_stAddress1 = m_rnFind1.Address
                    'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                    m_rnFind1.Offset(0, -4).Font.Color = vbRed
                    
                    
                    'MsgBox ("first" & m_stAddress)
                    
                    'Unhide the column, and then find the next X.
                        Do
                                                       
                            Set lastrowNonBorrowing1 = m_rnFind1
                            Set m_rnFind1 = .FindNext(m_rnFind1)
                            
                            'If m_rnFind.Offset(0, -10) = "N" Then
                            'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                            m_rnFind1.Offset(0, -4).Font.Color = vbRed
                            Counter1 = Counter1 + 1
                            'End If
                            
                             'MsgBox (m_rnFind)
                            'MsgBox ("m_rnFind.Address" & m_rnFind.Address)
                             'MsgBox ("Counter" & Counter)
                        Loop While Not m_rnFind1 Is Nothing And m_rnFind1.Address <> m_stAddress1 And Counter1 <> 19
                    End If
            End With
          
          
          Set m_rnCheck = Range("O" & 1 & ":" & "O" & NumRows)
          
          'If Not IsEmpty(Cells(iRow, "M")) Then
            'For Each Cell In Range("M" & StartduplicateNum & ":" & "M" & StartduplicateNum + Counter)
             'For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                'bln = False
            'MsgBox (NumRows)
              
                    
           'With m_rnCheck
             Set m_rnFind = Columns("O").Find(What:="Non-Borrowing Customers") 'LookIn:=xlFormulas)
             
                    
                If Not m_rnFind Is Nothing Then
                        strFirst = m_rnFind.Address
                    
                             
                Do
                
                
                If (Cells(m_rnFind.Row, "E").Text) = "N" Then
                'Found a match
           'MsgBox (m_rnFind.Address)
                m_stAddress = m_rnFind.Address
                m_rnFind.Offset(0, -14).Font.Color = vbBlue
                'MsgBox (m_rnFind.Address & Cells(m_rnFind.Row, "E").Value)
                Exit Do
                
                Else
                
                End If
                'Set m_rnFindnextfirst = Columns("O").Find(What:="Non-Borrowing Customers", m_rnFind.row, xlValues, xlWhole)
                Set m_rnFind = Columns("O").FindNext(m_rnFind)
                
           'MsgBox ("FINDnext" & m_rnFind.Address)
               
                Loop While m_rnFind.Address <> strFirst And (Cells(m_rnFind.Row, "E").Value) <> "N"
                m_rnFind.Offset(0, -14).Font.Color = vbBlue
                
                End If
                  
                  
                      'm_stAddress = m_rnFind.Address
                    
                     'm_rnFind.Offset(0, -14).Font.Color = vbBlue
                    
            
                        Do
                                                       
                            Set lastrowNonBorrowing = m_rnFind
                            Set m_rnFind = Columns("O").FindNext(m_rnFind)
                            
                            If m_rnFind.Offset(0, -10) = "N" Then
                            'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                            m_rnFind.Offset(0, -14).Font.Color = vbBlue
                            Counter = Counter + 1
                            End If
                            
                             'MsgBox (m_rnFind)
                            'MsgBox ("m_rnFind.Address" & m_rnFind.Address)
                             'MsgBox ("Counter" & Counter)
                        Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress And Counter <> 19
                    'End If
            
            'End With
            'MsgBox (Counter)
            If Counter < 19 Then
            MsgBox ("There are less than 20 Non-Borrowing Customers")
            End If
            
            'MsgBox ("last row" & lastrowNonBorrowing.Row)
            
     'Next
End Sub

Sub CheckFirst20NonBorrowingPart2()


Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet

Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls;*.xlsx), *.xls; *.xlsx", Title:="******Please choose your THIS MONTH CASA TOP 20 Workbook*****")
filename = Application.GetOpenFilename(Title:="******Please choose your THIS MONTH CASA TOP 20 (before add co) Workbook*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in your CASA TOP 20 Workbook
'Set mySheetName = myFileName.ActiveSheet
Set mySheetName = myFileName.Worksheets("Top20 Non-Borrowing Cust_Adjust")


'find last row in Column E ("PARTY_NUMBER")in your PREVIOUS MONTH CASA TOP 20 Workbook
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in your PREVIOUS MONTH CASA TOP 20 Workbook
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow  ' lastrowNonBorrowing.Row
    With currentSheet
        
        'If .Cells(i, "A").Interior.ColorIndex = 23 Then
         If .Cells(i, "A").Font.Color = vbBlue Or .Cells(i, "A").Font.Color = vbRed Then
         If Not IsError(Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)) Then
            matchRow = Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "J").Value
            '.Cells(i, "N").Interior.ColorIndex = 22 'pink
            '.Cells(i, "A").Interior.ColorIndex = 43 'green
            
         Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            '.Cells(i, "N") = CVErr(xlErrNA)
            .Cells(i, "A").Interior.ColorIndex = 43 'green
            
         End If
        Else 'nothing
        End If
      End With
Next i

myFileName.Close Save = False
End Sub

Sub CheckFirst20NonBorrowingSorting()
    Dim N As Long
    'If Intersect(Target, Range("A:A")) Is Nothing Then Exit Sub
    'N = Cells(Rows.Count, "A").End(xlUp).Row
    NumRows = Range("F1", Range("F1").End(xlDown)).Rows.Count
    ThisWorkbook.Sheets("CASASheet1").Range("A1:U" & NumRows).Sort Key1:=Range("F1"), Order1:=xlDescending, Header:=xlYes
End Sub


Sub Step6ThisMonthCASATop20changeCUSTTYPEnew1()


Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet

Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls), *.xls", Title:="******Please choose your THIS MONTH CASA TOP 20 Workbook*****")
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls;*.xlsx), *.xls; *.xlsx", Title:="******Please choose your THIS MONTH CASA TOP 20 Workbook*****")
filename = Application.GetOpenFilename(Title:="******Please choose your THIS MONTH CASA TOP 20 (before add co) Workbook*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in your CASA TOP 20 Workbook
'Set mySheetName = myFileName.ActiveSheet
Set mySheetName = myFileName.Worksheets("Top20 Non-Borrowing Cust_Adjust")

'find last row in Column A ("PARTY_NUMBER")in your CASA TOP 20 Workbook
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in your CASA TOP 20 Workbook
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in "CASASheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)) Then
            matchRow = Application.Match(Application.WorksheetFunction.Trim(.Cells(i, "A")), myRangeName, 0)
            .Cells(i, "N") = mySheetName.Cells(matchRow, "M").Value
            '.Cells(i, "N").Interior.ColorIndex = 22 'pink
            '.Cells(i, "A").Interior.ColorIndex = 22 'pink
            
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "N") = CVErr(xlErrNA)
            
        End If
        
        currentSheet.Cells(i, "O") = .Cells(i, "M")
        If IsError(currentSheet.Cells(i, "N").Value) Then
         
         Else
           
           If currentSheet.Cells(i, "M") = .Cells(i, "N") Then
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
           Else
                      currentSheet.Cells(i, "O") = .Cells(i, "N")
                      currentSheet.Cells(i, "O").Interior.ColorIndex = 17 ' purple
                      
           End If
        End If
    End With
Next i

myFileName.Close Save = False

End Sub

Sub AddTrimCIF()

 'Columns("G:H").EntireColumn.Delete
'Columns(8).EntireColumn.Delete
    
Range("B:B").EntireColumn.Insert
'Range("B:B").EntireColumn.Interior.Pattern = xlNone
'Range("B:B").EntireColumn.Font.ColorIndex = xlAutomatic
With ThisWorkbook.ActiveSheet
    
            '.Cells(1, 14) = "Customer Type Top20 Only"
            .Cells(1, 2) = "Trim CIF No."
    
    
    For Each mycell In ActiveSheet.Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "B") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
        .Cells(mycell.Row, "B").Interior.Pattern = xlNone
        .Cells(mycell.Row, "B").Font.ColorIndex = xlAutomatic
    Next mycell
        
    
    
End With

End Sub

Sub DatavalidationPersonal()

Dim MyList(2) As String
MyList(0) = "N"
MyList(1) = "Y"
MyList(2) = "#N/A"
'MyList(3) = 4
'MyList(4) = 5
'MyList(5) = 6

With Columns("E").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=Join(MyList, ",")
End With

End Sub

Sub DatavalidationCUSTTYPE()

Dim MyList(3) As String
MyList(0) = "Bills Customer"
MyList(1) = "Non-Bills Borrowing Customer"
MyList(2) = "Non-Borrowing Customers"
MyList(3) = "#N/A"
'MyList(4) = 5
'MyList(5) = 6

With Columns("L:O").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=Join(MyList, ",")
End With

End Sub


Sub ChangeDatavalidationCUSTTYPE()

With ThisWorkbook.Worksheets("CASAsheet1").UsedRange
    'With .Cells(1, 1).CurrentRegion
        'With .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)
            
            .Cells.Validation.Delete
End With


Dim MyList(2) As String
MyList(0) = "N"
MyList(1) = "Y"
MyList(2) = "#N/A"
'MyList(3) = 4
'MyList(4) = 5
'MyList(5) = 6

With Columns("F").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=Join(MyList, ",")
End With


Dim CustList(3) As String
CustList(0) = "Bills Customer"
CustList(1) = "Non-Bills Borrowing Customer"
CustList(2) = "Non-Borrowing Customers"
CustList(3) = "#N/A"
'MyList(4) = 5
'MyList(5) = 6

With Columns("M:P").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:=Join(CustList, ",")
End With

End Sub



Sub Section2Step1DSdirectGroupIDvlookup()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long
Dim Range1             As Range
Dim Range2             As Range
Dim crit1 As String
Dim crit2 As String

'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose THIS MONTH Cust_info_YYYYMM*****")
filename = Application.GetOpenFilename(Title:="******Please choose THIS MONTH Customer Enquiry_YYYYMM FINAL*****")

'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Customer Enquiry_YYYYMM FINAL
Set mySheetName = myFileName.Worksheets("DSDirectSheet1")

'find last row in Column E ("Trim CIF")in Customer Enquiry_YYYYMM FINAL
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in Customer Enquiry_YYYYMM FINAL
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

'set the range for Vlookup all active rows and columns in Customer Enquiry_YYYYMM FINAL
'Set Range1 = mySheetName.Range("G1:G" & lastrow)

'set the range for Vlookup all active rows and columns in Customer Enquiry_YYYYMM FINAL
'Set Range2 = mySheetName.Range("I1:I" & lastrow)

' find last row in Column B in This Workbook ("Trim CIF.") in "CASASheet1"
lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "B").End(xlUp).Row
'MsgBox (lastRow)

'crit1 = "Active"
'crit2 = "DAVID LOOK"


For i = 2 To lastrow1

   With currentSheet
        'If Not IsError(Application.Match(.Cells(i, "B") , myRangeName, 0)) And mySheetName.Cells(i, "G") <> "Closed" And mySheetName.Cells(i, "I") <> CVErr(xlErrNA) Then
        'If Not IsError(Application.Match(.Cells(i, "B") & crit1 & crit2, myRangeName & Range1 & Range2, 0)) Then
          If Not IsError(Application.Match(.Cells(i, "B"), myRangeName, 0)) Then
            
            matchRow = Application.Match(.Cells(i, "B"), myRangeName, 0)
                'matchRow = Application.Match(.Cells(i, "B") & crit1 & crit2, myRangeName & Range1 & Range2, 0)
            'If mySheetName.Cells(matchRow, "G").Value <> "Closed" And mySheetName.Cells(matchRow, "I").Value <> CVErr(xlErrNA) Then
            
            If mySheetName.Cells(matchRow, "G").Value <> "Closed" Then
            
                currentSheet.Cells(i, "W") = mySheetName.Cells(matchRow, "A").Value
            
            Else
                currentSheet.Cells(i, "W") = CVErr(xlErrNA)
            
            End If
            '.Cells(i, "H") = mySheetName.Cells(matchRow, "D").Value
            '.Cells(i, "I") = mySheetName.Cells(matchRow, "E").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "G").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            currentSheet.Cells(i, "W") = CVErr(xlErrNA)
            '.Cells(i, "H") = CVErr(xlErrNA)
            '.Cells(i, "I") = CVErr(xlErrNA)
            '.Cells(i, "J") = CVErr(xlErrNA)
            '.Cells(i, "K") = CVErr(xlErrNA)
        End If

    End With
Next i


myFileName.Close Save = False

End Sub


Sub Section2Step2PercentageOfCASAofEachZone()

Dim x As Integer
Dim a As Integer
Dim NumRows As Integer
Dim SumOfAvgAmount As Double

      
   Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")
   Set Mysheets = ThisWorkbook.Worksheets("DataValidation")
   
   For B = 2 To Mysheets.Cells(Mysheets.Rows.Count, "E").End(xlUp).Row
    
   EachZoneSumOfAvgAmount = 0
    NumRows = currentSheet.Range("A1", Range("A1").End(xlDown)).Rows.Count
    For a = 2 To NumRows
     'If currentSheet.Range("L" & a).Value = Mysheets.Cells(b, "E").Value Then
     If WorksheetFunction.Trim(currentSheet.Range("L" & a)) = WorksheetFunction.Trim(Mysheets.Range("E" & B)) Then
     
        
    EachZoneSumOfAvgAmount = EachZoneSumOfAvgAmount + currentSheet.Range("G" & a).Value
     End If
    Next
    'MsgBox (SumOfAvgAmount)
    
    
    For a = 2 To NumRows
    
    'If currentSheet.Range("L" & a).Value = Mysheets.Cells(b, "E").Value Then
    If WorksheetFunction.Trim(currentSheet.Range("L" & a)) = WorksheetFunction.Trim(Mysheets.Range("E" & B)) Then
    
    For Each Cell In Range("X" & a & ":" & "X" & a)
                    Cell.Value = Cell.Offset(0, -17) / EachZoneSumOfAvgAmount
                    Cell.NumberFormat = "0.00%"
                    
    Next
    End If
    Next
    
    
   Next
End Sub

Sub Section2step3DividedofTop20Fils()

Dim Wb As ThisWorkbook
Dim ws As Worksheets
Dim Mysheets As Worksheet
Dim B As Integer

Set Mysheets = ThisWorkbook.Worksheets("DataValidation")
   
 For B = 2 To Mysheets.Cells(Mysheets.Rows.Count, "E").End(xlUp).Row

    Application.DisplayAlerts = False

    Rem Copy Data From NRM_Homing_Upload
    With ThisWorkbook.Sheets("CASASheet1")

    Dim lRow As Long
    lRow = .Range("A" & .Rows.Count).End(xlUp).Row

    With .Range("A1:X" & lRow)

        '.AutoFilter 12, "=BERNARD LEE"
        
        .AutoFilter 12, Mysheets.Range("E" & B).Value
        .AutoFilter 6, "=N"
        'CopyToNewBook ThisWorkbook, ThisWorkbook.Sheets("CASASheet1"), .SpecialCells(xlCellTypeVisible), "Regional Top 20 Corp Cust List " & Year(Date) & Format(DateAdd("m", -1, Now), "MM") & " BERNARD LEE"
        
        .SpecialCells(xlCellTypeVisible).Copy
        CopyToNewBook ThisWorkbook, ThisWorkbook.Sheets("CASASheet1"), .SpecialCells(xlCellTypeVisible), "Regional Top 20 Corp Cust List " & Year(Date) & Format(DateAdd("m", -1, Now), "MM") & " " & Mysheets.Range("E" & B)
        
        
        '.AutoFilter 1, "<>*001"

        'CopyToNewBook ThisWorkbook, ThisWorkbook.Sheets("NRM_Homing_Upload"), .SpecialCells(xlCellTypeVisible), "LC"

    End With

    .AutoFilterMode = False

    End With

Next



End Sub

Sub CopyToNewBook(Wb As Workbook, ws As Worksheet, rng As Range, sFile As String)

Dim new_book As Workbook
Dim i As Integer
Dim j As Integer

'Wb.Sheets(ws.Name).Range(rng.Address).Copy

'ThisWorkbook.Worksheets("CASAsheet1").Range(rng.Address).Copy

'Worksheets("CASAsheet1").SpecialCells(xlCellTypeVisible).Copy


Set new_book = Workbooks.Add
With new_book

    With .Sheets(1)
        
        'wb.Sheets(ws.Name).Range(rng.Address).Copy
        
           
        Dim ARow As Long
        ARow = new_book.Sheets(1).Range("A" & .Rows.Count).End(xlDown).Row
        
        '.Range("a1").PasteSpecial (xlPasteAll)
        .Range("a1").PasteSpecial Paste:=xlPasteValues
        
        '.UsedRange.RemoveDuplicates Columns:=8, Header:=xlYes
        .Range("A22:X" & ARow).Delete
        
                
        .Columns("B:F").clearcontents
        .Range("B1:B21").Value = Range("H1:H21").Value
        .Range("C1:C21").Value = Range("I1:I21").Value
        .Columns("H:I").clearcontents
        
        .Range("D1:D21").Value = Range("G1:G21").Value
        .Columns("G").clearcontents
        
        .Range("E1:E21").Value = Range("L1:L21").Value
        .Columns("L").clearcontents
        
        .Range("F1:F21").Value = Range("J1:J21").Value
        .Range("G1:G21").Value = Range("K1:K21").Value
        .Columns("J:K").clearcontents
        
        .Range("H1:H21").Value = Range("P1:P21").Value
        .Columns("M:P").clearcontents
        
        .Range("I1:I21").Value = Range("V1:V21").Value
        .Columns("V").clearcontents
        
        .Range("J1:J21").Value = Range("W1:W21").Value
        .Columns("W").clearcontents
        
        .Range("K1:K21").Value = Range("Q1:Q21").Value
        .Columns("Q").clearcontents
        
        .Range("L1:L21").Value = Range("X1:X21").Value
        .Columns("X").clearcontents
        .Columns("R:U").clearcontents
        .UsedRange.Columns.AutoFit
        
        
        .Range("A1") = "BR No"
        .Range("B1") = "BR name"
        .Range("C1") = "Business Nature"
        .Range("D1") = "MTH_AVG"
        .Range("E1") = "Region"
        .Range("F1") = "OFFICER_CD"
        .Range("G1") = "OFFICER_SUB_CD"
        .Range("H1") = "CUST TYPE(Final)"
        .Range("I1") = "Tier Int"
        .Range("J1") = "DS-Direct"
        .Range("K1") = "Penetration Ratio on total CBD balance"
        .Range("L1") = "Penetration Ratio on Regional CASA balance"
                
        .Rows("1").Font.Bold = True
        .Rows("1").WrapText = True
        
        .Columns("D").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
        .Columns("K:L").NumberFormat = "0.00%"
        
                   
        For i = 2 To 21
        
        If IsError(Range("I" & i).Value) Then
            If Range("I" & i).Value = CVErr(xlErrNA) Then
            Range("I" & i).Value = "N"
            End If
        Else: Range("I" & i).Value = "Y"
        End If
        
        If IsError(Range("J" & i).Value) Then
            If Range("J" & i).Value = CVErr(xlErrNA) Then
            Range("J" & i).Value = "N"
            End If
        Else: Range("J" & i).Value = "Y"
        End If
        
        
        Next
        
                          
    End With




    .SaveAs filename:=ThisWorkbook.Path & "\" & sFile & ".xlsx"
    '.SaveAs Filename:=ThisWorkbook.Path & " \testing.xlsx"
    .Close

End With

End Sub

Sub Section2step4TotalofTop20Fils()

Dim Wb As ThisWorkbook
Dim ws As Worksheets
Dim Mysheets As Worksheet
Dim B As Integer

 
 Set new_bookTotal = Workbooks.Add
 
 With new_bookTotal
    With .Sheets(1)
 .Range("A1") = "BR No"
        .Range("B1") = "BR name"
        .Range("C1") = "Business Nature"
        .Range("D1") = "MTH_AVG"
        .Range("E1") = "Region"
        .Range("F1") = "OFFICER_CD"
        .Range("G1") = "OFFICER_SUB_CD"
        .Range("H1") = "CUST TYPE(Final)"
        .Range("I1") = "Tier Int"
        .Range("J1") = "DS-Direct"
        .Range("K1") = "Penetration Ratio on total CBD balance"
        .Range("L1") = "Penetration Ratio on Regional CASA balance"
                
        .Rows("1").Font.Bold = True
        .Rows("1").WrapText = True
        
     End With

  
 Set Mysheets = ThisWorkbook.Worksheets("DataValidation")
 
 For B = 2 To Mysheets.Cells(Mysheets.Rows.Count, "E").End(xlUp).Row

    Application.DisplayAlerts = False

    Rem Copy Data From NRM_Homing_Upload
    With ThisWorkbook.Sheets("CASASheet1")

    Dim lRow As Long
    lRow = .Range("A" & .Rows.Count).End(xlUp).Row

    With .Range("A1:X" & lRow)

        '.AutoFilter 12, "=BERNARD LEE"
        
        .AutoFilter 12, Mysheets.Range("E" & B).Value
        .AutoFilter 6, "=N"
        'CopyToNewBook ThisWorkbook, ThisWorkbook.Sheets("CASASheet1"), .SpecialCells(xlCellTypeVisible), "Regional Top 20 Corp Cust List " & Year(Date) & Format(DateAdd("m", -1, Now), "MM") & " BERNARD LEE"
        
        .SpecialCells(xlCellTypeVisible).Copy
        'CopyToNewBook1Testing ThisWorkbook, ThisWorkbook.Sheets("CASASheet1"), .SpecialCells(xlCellTypeVisible), "Regional Top 20 Corp Cust List " & Year(Date) & Format(DateAdd("m", -1, Now), "MM") & " " & Mysheets.Range("E" & B)
        
        
        
        
        With new_bookTotal.Sheets(1)
    ''With new_bookTotal
        'wb.Sheets(ws.Name).Range(rng.Address).Copy
        '''With .Sheets(1)
        'Dim ARow As Long
        ARow = new_bookTotal.Sheets(1).Range("A" & .Rows.Count).End(xlUp).Row
        'MsgBox (.Range("A" & .Rows.Count).End(xlDown).Row)
        c = ARow
        'MsgBox ("c=" & c)
        
        
        '.Range("a1").PasteSpecial (xlPasteAll)
        '''.Range("a1").PasteSpecial Paste:=xlPasteValues
        new_bookTotal.Sheets(1).Range("A" & ARow + 1).PasteSpecial Paste:=xlPasteValues
        new_bookTotal.Sheets(1).Range("A" & c + 1 & ":X" & c + 1).Delete
        
        
        BROW = new_bookTotal.Sheets(1).Range("A" & .Rows.Count).End(xlUp).Row
        'MsgBox ("BROW=" & BROW)
        
        '.UsedRange.RemoveDuplicates Columns:=8, Header:=xlYes
        new_bookTotal.Sheets(1).Range("A" & c + 21 & ":X" & BROW).Delete
        
        End With
     
    End With

    .AutoFilterMode = False

    End With

Next



 With new_bookTotal.Sheets(1)


CRow = new_bookTotal.Sheets(1).Range("A" & .Rows.Count).End(xlUp).Row

 .Columns("B:F").clearcontents
        .Range("B1:B" & CRow).Value = Range("H1:H" & CRow).Value
        .Range("C1:C" & CRow).Value = Range("I1:I" & CRow).Value
        .Columns("H:I").clearcontents
        
        .Range("D1:D" & CRow).Value = Range("G1:G" & CRow).Value
        .Columns("G").clearcontents
        
        .Range("E1:E" & CRow).Value = Range("L1:L" & CRow).Value
        .Columns("L").clearcontents
        
        .Range("F1:F" & CRow).Value = Range("J1:J" & CRow).Value
        .Range("G1:G" & CRow).Value = Range("K1:K" & CRow).Value
        .Columns("J:K").clearcontents
        
        .Range("H1:H" & CRow).Value = Range("P1:P" & CRow).Value
        .Columns("M:P").clearcontents
        
        .Range("I1:I" & CRow).Value = Range("V1:V" & CRow).Value
        .Columns("V").clearcontents
        
        .Range("J1:J" & CRow).Value = Range("W1:W" & CRow).Value
        .Columns("W").clearcontents
        
        .Range("K1:K" & CRow).Value = Range("Q1:Q" & CRow).Value
        .Columns("Q").clearcontents
        
        .Range("L1:L" & CRow).Value = Range("X1:X" & CRow).Value
        .Columns("X").clearcontents
        .Columns("R:U").clearcontents
        .UsedRange.Columns.AutoFit
        
        
        .Columns("D").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
        .Columns("K:L").NumberFormat = "0.00%"
         
         
         
        .Range("A1") = "BR No"
        .Range("B1") = "BR name"
        .Range("C1") = "Business Nature"
        .Range("D1") = "MTH_AVG"
        .Range("E1") = "Region"
        .Range("F1") = "OFFICER_CD"
        .Range("G1") = "OFFICER_SUB_CD"
        .Range("H1") = "CUST TYPE(Final)"
        .Range("I1") = "Tier Int"
        .Range("J1") = "DS-Direct"
        .Range("K1") = "Penetration Ratio on total CBD balance"
        .Range("L1") = "Penetration Ratio on Regional CASA balance"
                
        .Rows("1").Font.Bold = True
        .Rows("1").WrapText = True
                          
                          
          For i = 2 To CRow
        
        If IsError(Range("I" & i).Value) Then
            If Range("I" & i).Value = CVErr(xlErrNA) Then
            Range("I" & i).Value = "N"
            End If
        Else: Range("I" & i).Value = "Y"
        End If
        
        If IsError(Range("J" & i).Value) Then
            If Range("J" & i).Value = CVErr(xlErrNA) Then
            Range("J" & i).Value = "N"
            End If
        Else: Range("J" & i).Value = "Y"
        End If
       Next


End With




.SaveAs filename:=ThisWorkbook.Path & "\" & "Regional Top 20 Corp Cust List " & Year(Date) & Format(DateAdd("m", -1, Now), "MM") & "- All Zone" & ".xlsx"
.Close
End With

End Sub
