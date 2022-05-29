Sub Step1clearcontents()

With ThisWorkbook.Worksheets("DSDirectsheet1").UsedRange
    'With .Cells(1, 1).CurrentRegion
        'With .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)
            .Cells.clearcontents
            .Cells.Interior.Pattern = xlNone
            .Cells.Font.ColorIndex = xlAutomatic
            '.Cells.Validation.Delete
            '.Columns("I").NumberFormat = "General"
            '.Columns("J").NumberFormat = "General"
            '.Cells(1, 1) = "PARTY_NUMBER"
            '.Cells(1, 2) = "ID Indicator(Personal = ID, CI, PP, ZZ, OTW,Company= BR, OT, XX)"
            '.Cells(1, 3) = "PARTY_NUMBER(First 10 to unlimited digits)"
            '.Cells(1, 4) = "Suffix(Left 3 digits for BR company)"
            '.Cells(1, 5) = "Personal?"
            '.Cells(1, 6) = "monthly_average"
            '.Cells(1, 7) = "FULL_NAME"
            '.Cells(1, 8) = "OCCUPATION"
            '.Cells(1, 9) = "OFFICER_CD"
            '.Cells(1, 10) = "OFFICER_SUB_CD"
            '.Cells(1, 11) = "RH"
            
        'End With
    'End With
End With


Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "DSDirectSheet1" Then
    
    Else: ws.Delete
    End If
Next
Application.DisplayAlerts = True


End Sub
Sub Step2importfiles()

Call S2ImportCustomerEnquiry_YYYYMM
Call S3AddTitle
Call S4DSdirectVlookupTHisMonthCASAInfo
Call S5DSdirectvlookupPreviousCASAInfo
Call S6DSdirectvlookupLastMonthCustomerEnquiry_YYYYMMDD
Call S7DSdirectvlookupDSESDSHKOthersDSDCust
Call S8importCurrentMonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM
Call S9indicateDSHK00001andMTR
Call s10IndicateMTRmaster
Call S11importCurrent_minus_1_MonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM
Call S12importCurrent_minus_2_MonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM
Call S13threeMonthsInARow
Call S14ExistingClientAsOfDec2020
Call S15matchRMName

End Sub

Sub S2ImportCustomerEnquiry_YYYYMM()

Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Sheet As Worksheet
Dim PasteStart As Range

Set wb1 = ActiveWorkbook
Set PasteStart = [DSDirectSheet1!A1]

FileToOpen = Application.GetOpenFilename _
(Title:="******Please choose THIS MONTH CustomerEnquiry_YYYYMM*****", _
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
            .Offset(0, 1).EntireColumn.Copy PasteStart.Range("B1")
            .Offset(0, 2).EntireColumn.Copy PasteStart.Range("C1")
            .Offset(0, 3).EntireColumn.Copy PasteStart.Range("D1")
            .Offset(0, 4).EntireColumn.Copy PasteStart.Range("E1")
            .Offset(0, 5).EntireColumn.Copy PasteStart.Range("F1")
            .Offset(0, 6).EntireColumn.Copy PasteStart.Range("G1")
            .Offset(0, 7).EntireColumn.Copy PasteStart.Range("H1")
        End With
    Next Sheet
End If

    wb2.Close saveChanges:=False
    
End Sub

Sub S3AddTitle()

 Columns("G:H").EntireColumn.Delete
'Columns(8).EntireColumn.Delete
    
Range("E:E").EntireColumn.Insert

With ThisWorkbook.ActiveSheet
    
            
    .Cells(1, 5) = "TRIM CIF No."
    .Cells(1, 5).Interior.ColorIndex = 6 'yellow
    .Cells(1, 5).WrapText = True
    
    .Cells(1, 7) = "Company Status DO NOT SELECT CLOSED"
    .Cells(1, 7).Interior.ColorIndex = 6 'yellow
    .Cells(1, 7).WrapText = True
    
    .Cells(1, 8) = "Customer Type from CASA FINAL"
    .Cells(1, 8).WrapText = True
    .Cells(1, 9) = "Zone from CASA FINAL"
    .Cells(1, 9).WrapText = True
    
    .Cells(1, 10) = "OFFICER_CD"
    .Cells(1, 11) = "OFFICER_SUB_CD"
                
    .Cells(1, 12) = "Login Status FROM ACTIVE LIST"
    .Cells(1, 12).Interior.ColorIndex = 6 'yellow
    .Cells(1, 12).WrapText = True
    
    .Cells(1, 13) = "Payment Customer from Active list"
    .Cells(1, 13).Interior.ColorIndex = 6 'yellow
    .Cells(1, 13).WrapText = True
    
    .Cells(1, 14) = "Trade Customer from Active list"
    .Cells(1, 14).Interior.ColorIndex = 6 'yellow
    .Cells(1, 14).WrapText = True
    
    .Cells(1, 15) = "MTR DSHK0700 and DSHK00001 DO NOT SELECT Y"
    .Cells(1, 15).Interior.ColorIndex = 6 'yellow
    .Cells(1, 15).WrapText = True
    
            '.Cells(1, 14) = "Customer Type Top20 Only"
            '.Cells(1, 15) = "Customer Type Final"
    
    .Cells(1, 16) = "DS-direct Open Date"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 16).WrapText = True
    
    .Cells(1, 17) = "Login Status FROM ACTIVE LIST (Current-1)"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 17).WrapText = True
    
    .Cells(1, 18) = "Login Status FROM ACTIVE LIST (Current-2)"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 18).WrapText = True
    
    
    .Cells(1, 19) = "Active 3 months in a row?"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 19).WrapText = True
    
    .Cells(1, 20) = "Existing or New Client as of Dec 2020?"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 20).WrapText = True
    
     .Cells(1, 21) = "AO Name"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 21).WrapText = True
    
     .Cells(1, 22) = "Team Head Name"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 22).WrapText = True
    
     .Cells(1, 23) = "RH Name"
    '.Cells(1, 16).Interior.ColorIndex = 6 'yellow
    .Cells(1, 23).WrapText = True
    
    For Each mycell In ActiveSheet.Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "E") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
        
        
        
    
    
End With

End Sub


Sub S4DSdirectVlookupTHisMonthCASAInfo()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
'Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")
Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your CASA Final file*****")
filename = Application.GetOpenFilename(Title:="******Please choose your Current Month CASA info*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in CASA info
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("CASASheet1")

'find last row in Column A ("PARTY_NUMBER")in CASA info
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in CASA info
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(Trim(.Cells(i, "D")), myRangeName, 0)) Then
            matchRow = Application.Match(Trim(.Cells(i, "D")), myRangeName, 0)
            .Cells(i, "I") = mySheetName.Cells(matchRow, "G").Value
            '.Cells(i, "I") = mySheetName.Cells(matchRow, "L").Value
            .Cells(i, "J") = mySheetName.Cells(matchRow, "E").Value
            .Cells(i, "K") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "G").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "I") = CVErr(xlErrNA)
            '.Cells(i, "I") = CVErr(xlErrNA)
            .Cells(i, "J") = CVErr(xlErrNA)
            .Cells(i, "K") = CVErr(xlErrNA)
            '.Cells(i, "K") = CVErr(xlErrNA)
        End If

    End With
Next i


myFileName.Close saveChanges:=False

End Sub

Sub S5DSdirectvlookupPreviousCASAInfo()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
'Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")
Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your Previous Month CASA Final file*****")
filename = Application.GetOpenFilename(Title:="******Please choose your Previous Month CASA info*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in CASA Final File
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("CASASheet1")

'find last row in Column A ("PARTY_NUMBER")in CASA info
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in CASA info
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        
        If IsError(currentSheet.Cells(i, "I").Value) Then
        
        If Not IsError(Application.Match(Trim(.Cells(i, "D")), myRangeName, 0)) Then
            matchRow = Application.Match(Trim(.Cells(i, "D")), myRangeName, 0)
            .Cells(i, "I") = mySheetName.Cells(matchRow, "G").Value
            .Cells(i, "I").Interior.ColorIndex = 8 'Cyan
            '.Cells(i, "I") = mySheetName.Cells(matchRow, "L").Value
            '.Cells(i, "I").Interior.ColorIndex = 8 'Cyan
            .Cells(i, "J") = mySheetName.Cells(matchRow, "E").Value
            .Cells(i, "J").Interior.ColorIndex = 8 'Cyan
            .Cells(i, "K") = mySheetName.Cells(matchRow, "F").Value
            .Cells(i, "K").Interior.ColorIndex = 8 'Cyan
            
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


Sub S6DSdirectvlookupLastMonthCustomerEnquiry_YYYYMMDD()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Curent month Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DSES+DSHK - Others (DS-D Cust)*****")
filename = Application.GetOpenFilename(Title:="******Please choose your previous month Customer Enquiry FINAL*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in previous month Customer Enquiry FINAL
'Set mySheetName = myFileName.ActiveSheet
Set mySheetName = myFileName.Worksheets("DSDirectSheet1")

'find last row in Column D ("PARTY_NUMBER")in previous month Customer Enquiry FINAL
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in previous month Customer Enquiry FINAL
Set myRangeName = mySheetName.Range("E1:E" & lastrow)

' find last row in Column D in This Workbook ("PARTY_NUMBER.") in activesheet in Current Month Customer Enquiry_YYYYMMDD
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "E").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        
        If IsError(currentSheet.Cells(i, "I").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            '.Cells(i, "H") = mySheetName.Cells(matchRow, "H").Value
            '.Cells(i, "H").Interior.ColorIndex = 4 'Green
            .Cells(i, "I") = mySheetName.Cells(matchRow, "I").Value
            .Cells(i, "I").Interior.ColorIndex = 4 'Green
            .Cells(i, "J") = mySheetName.Cells(matchRow, "J").Value
            .Cells(i, "J").Interior.ColorIndex = 4 'Green
            .Cells(i, "K") = mySheetName.Cells(matchRow, "K").Value
            .Cells(i, "K").Interior.ColorIndex = 4 'Green
            
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

Sub S7DSdirectvlookupDSESDSHKOthersDSDCust()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
'Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")
Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DSES+DSHK - Others (DS-D Cust)*****")
filename = Application.GetOpenFilename(Title:="******Please choose your DSES+DSHK - Others (DS-D Cust)*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in DSES+DSHK - Others (DS-D Cust)
'Set mySheetName = myFileName.ActiveSheet
Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")

'find last row in Column A ("PARTY_NUMBER")in DSES+DSHK - Others (DS-D Cust)
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "I").End(xlUp).Row

'set the range for Vlookup all active rows and columns in DSES+DSHK - Others (DS-D Cust)
Set myRangeName = mySheetName.Range("I1:I" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        
        If IsError(currentSheet.Cells(i, "I").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            '.Cells(i, "H") = mySheetName.Cells(matchRow, "K").Value
            '.Cells(i, "H").Interior.ColorIndex = 7 'Magenta
            .Cells(i, "I") = mySheetName.Cells(matchRow, "J").Value
            .Cells(i, "I").Interior.ColorIndex = 7 'Magenta
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
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

Sub S8importCurrentMonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DS_CBD_TB_ACT_RPT21_active_list_201912*****")
filename = Application.GetOpenFilename(Title:="******Please choose your Current Month DS_CBD_TB_ACT_RPT21_active_list_YYYYMM*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Active List
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")


mySheetName.Range("E:E").EntireColumn.Insert

'With ThisWorkbook.mySheetName
 With mySheetName
            
    .Cells(1, 5) = "TRIM CIF No."
    .Cells(1, 5).Interior.ColorIndex = 6 'yellow
    .Cells(1, 5).WrapText = True
  
    
    For Each mycell In mySheetName.Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "E") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
  
End With

'myFileName.Save

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
'find last row in Column A in This Workbook ("PARTY_NUMBER.") in "DSDirectSheet1" in DSdirect Macro

'find last row in Column A ("PARTY_NUMBER")in DSES+DSHK - Others (DS-D Cust)
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in DSES+DSHK - Others (DS-D Cust)
Set myRangeName = mySheetName.Range("E1:E" & lastrow)


lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        'If IsError(currentSheet.Cells(i, "H").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            .Cells(i, "L") = mySheetName.Cells(matchRow, "S").Value
            .Cells(i, "M") = mySheetName.Cells(matchRow, "H").Value
            .Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            .Cells(i, "P") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            
            If .Cells(i, "G") <> "Closed" Then
            
            .Cells(i, "L") = "Inactive(Unknown)"
            .Cells(i, "M") = "N (Unknown)"
            .Cells(i, "N") = "N (Unknown)"
            
            Else
            .Cells(i, "L") = CVErr(xlErrNA)
            .Cells(i, "M") = CVErr(xlErrNA)
            .Cells(i, "N") = CVErr(xlErrNA)
            .Cells(i, "P") = CVErr(xlErrNA)
            
            '.Cells(i, "L").Interior.ColorIndex = 8
            End If
        End If
        
        'End If
    End With
Next i


myFileName.Close saveChanges:=False

End Sub

Sub S9indicateDSHK00001andMTR()

Dim x As Integer
Dim a As Integer
Dim mycell As Range
Dim NumRows As Integer

    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'MsgBox (NumRows)
    For a = 2 To NumRows


   For Each mycell In Range("A" & a & ":" & "A" & a)
    'If Cells(a, "A").Value = "DSHK00001" Or "DSHK0700" Then
    
    If mycell.Value = "DSHK00001" Or mycell.Value = "DSHK0700" Then
    
       mycell.Offset(0, 14).Value = "Y"
      
          
    End If
      
   Next
   Next
     
     
   
    
    
 
End Sub

Sub s10IndicateMTRmaster()

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
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    'For a = 2 To NumRows
    'MsgBox (NumRows)
       
           
          
          Set m_rnCheck1 = Range("A" & 1 & ":" & "A" & NumRows)
          
          With m_rnCheck1
             Set m_rnFind1 = .Find(What:="DSHK0700") 'LookIn:=xlFormulas)
                    If Not m_rnFind1 Is Nothing Then
                    m_stAddress1 = m_rnFind1.Address
                    'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                    m_rnFind1.Offset(0, 14).Value = "Y(Master)"
                    
                    
                    'MsgBox ("first" & m_stAddress)
                    
                    'Unhide the column, and then find the next X.
                        ''Do
                                                       
                            ''Set lastrowNonBorrowing1 = m_rnFind1
                            ''Set m_rnFind1 = .FindNext(m_rnFind1)
                            
                            'If m_rnFind.Offset(0, -10) = "N" Then
                            'm_rnFind.Offset(0, -14).Interior.ColorIndex = 23 'darkblue
                            ''m_rnFind1.Offset(0, -4).Font.Color = vbRed
                            ''Counter1 = Counter1 + 1
                            'End If
                            
                             'MsgBox (m_rnFind)
                            'MsgBox ("m_rnFind.Address" & m_rnFind.Address)
                             'MsgBox ("Counter" & Counter)
                        ''Loop While Not m_rnFind1 Is Nothing And m_rnFind1.Address <> m_stAddress1 And Counter1 <> 19
                    End If
            End With

End Sub

Sub S11importCurrent_minus_1_MonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DS_CBD_TB_ACT_RPT21_active_list_201912*****")
filename = Application.GetOpenFilename(Title:="******Please choose your (Current-1) Month DS_CBD_TB_ACT_RPT21_active_list_YYYYMM*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Active List
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")


mySheetName.Range("E:E").EntireColumn.Insert

'With ThisWorkbook.mySheetName
 With mySheetName
            
    .Cells(1, 5) = "TRIM CIF No."
    .Cells(1, 5).Interior.ColorIndex = 6 'yellow
    .Cells(1, 5).WrapText = True
  
    
    For Each mycell In mySheetName.Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "E") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
  
End With

'myFileName.Save

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
'find last row in Column A in This Workbook ("PARTY_NUMBER.") in "DSDirectSheet1" in DSdirect Macro

'find last row in Column A ("PARTY_NUMBER")in DSES+DSHK - Others (DS-D Cust)
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in DSES+DSHK - Others (DS-D Cust)
Set myRangeName = mySheetName.Range("E1:E" & lastrow)


lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        'If IsError(currentSheet.Cells(i, "H").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            
            .Cells(i, "Q") = mySheetName.Cells(matchRow, "S").Value
            '.Cells(i, "L") = mySheetName.Cells(matchRow, "S").Value
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "H").Value
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            
            If .Cells(i, "G") <> "Closed" Then
            
            .Cells(i, "Q") = "Inactive(Unknown)"
            '.Cells(i, "M") = "N (Unknown)"
            '.Cells(i, "N") = "N (Unknown)"
            
            Else
            .Cells(i, "Q") = CVErr(xlErrNA)
            '.Cells(i, "M") = CVErr(xlErrNA)
            '.Cells(i, "N") = CVErr(xlErrNA)
            '.Cells(i, "P") = CVErr(xlErrNA)
            
            '.Cells(i, "L").Interior.ColorIndex = 8
            End If
        End If
        
        'End If
    End With
Next i


myFileName.Close saveChanges:=False

End Sub

Sub S12importCurrent_minus_2_MonthDS_CBD_TB_ACT_RPT21_active_list_YYYYMM()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DS_CBD_TB_ACT_RPT21_active_list_201912*****")
filename = Application.GetOpenFilename(Title:="******Please choose your (Current-2) Month DS_CBD_TB_ACT_RPT21_active_list_YYYYMM*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Active List
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")


mySheetName.Range("E:E").EntireColumn.Insert

'With ThisWorkbook.mySheetName
 With mySheetName
            
    .Cells(1, 5) = "TRIM CIF No."
    .Cells(1, 5).Interior.ColorIndex = 6 'yellow
    .Cells(1, 5).WrapText = True
  
    
    For Each mycell In mySheetName.Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "E") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
  
End With

'myFileName.Save

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
'find last row in Column A in This Workbook ("PARTY_NUMBER.") in "DSDirectSheet1" in DSdirect Macro

'find last row in Column A ("PARTY_NUMBER")in DSES+DSHK - Others (DS-D Cust)
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "E").End(xlUp).Row

'set the range for Vlookup all active rows and columns in DSES+DSHK - Others (DS-D Cust)
Set myRangeName = mySheetName.Range("E1:E" & lastrow)


lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        'If IsError(currentSheet.Cells(i, "H").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            
            .Cells(i, "R") = mySheetName.Cells(matchRow, "S").Value
            '.Cells(i, "L") = mySheetName.Cells(matchRow, "S").Value
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "H").Value
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            
            If .Cells(i, "G") <> "Closed" Then
            
            .Cells(i, "R") = "Inactive(Unknown)"
            '.Cells(i, "M") = "N (Unknown)"
            '.Cells(i, "N") = "N (Unknown)"
            
            Else
            .Cells(i, "R") = CVErr(xlErrNA)
            '.Cells(i, "M") = CVErr(xlErrNA)
            '.Cells(i, "N") = CVErr(xlErrNA)
            '.Cells(i, "P") = CVErr(xlErrNA)
            
            '.Cells(i, "L").Interior.ColorIndex = 8
            End If
        End If
        
        'End If
    End With
Next i


myFileName.Close saveChanges:=False

End Sub

Sub S13threeMonthsInARow()

'Dim filename                As String
'Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
'Dim mySheetName             As Worksheet
'Dim myRangeName             As Range
'Dim lastrow                 As Long
Dim i                       As Long
'Dim matchRow                As Long
Dim lastrow1                  As Long

Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        If IsError(currentSheet.Cells(i, "L").Value) Or IsError(currentSheet.Cells(i, "Q").Value) Or IsError(currentSheet.Cells(i, "R").Value) Then
        
        .Cells(i, "S") = "N"
                    
        Else
        
            If .Cells(i, "L") = "Active" And .Cells(i, "Q") = "Active" And .Cells(i, "R") = "Active" Then
            .Cells(i, "S") = "Y"
            Else
            .Cells(i, "S") = "N"
            End If
        
        End If
               
        
    End With
Next i



End Sub

Sub S14ExistingClientAsOfDec2020()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DS_CBD_TB_ACT_RPT21_active_list_201912*****")
filename = Application.GetOpenFilename(Title:="******Please choose your BR List as of Dec 2020*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in BR list as of Dec 2020
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")


'mySheetName.Range("E:E").EntireColumn.Insert
mySheetName.Range("B:B").EntireColumn.Insert

'With ThisWorkbook.mySheetName
 With mySheetName
            
    .Cells(1, 2) = "TRIM CIF No."
    .Cells(1, 2).Interior.ColorIndex = 6 'yellow
    .Cells(1, 2).WrapText = True
  
    
    For Each mycell In mySheetName.Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "B") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
  
End With

'myFileName.Save

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
'find last row in Column A in This Workbook ("PARTY_NUMBER.") in "DSDirectSheet1" in DSdirect Macro
'find last row in Column A ("PARTY_NUMBER")in DSES+DSHK - Others (DS-D Cust)

'find last row in Column A ("PARTY_NUMBER")in Br List as of Dec 2020
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "B").End(xlUp).Row

'set the range for Vlookup all active rows and columns in DSES+DSHK - Others (DS-D Cust)

'set the range for Vlookup all active rows and columns in Br List as of Dec 2020
Set myRangeName = mySheetName.Range("B1:B" & lastrow)


lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        'If IsError(currentSheet.Cells(i, "H").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            
            .Cells(i, "T").Value = "Existing"
            '.Cells(i, "L") = mySheetName.Cells(matchRow, "S").Value
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "H").Value
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            
            .Cells(i, "T").Value = "New"
            
        End If
        
        'End If
    End With
Next i


myFileName.Close saveChanges:=False

End Sub

Sub S15matchRMName()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Customer Enquiry_YYYYMMDD
Set currentSheet = ThisWorkbook.Worksheets("DSDirectSheet1")
'Set currentSheet = ThisWorkbook.ActiveSheet

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your DS_CBD_TB_ACT_RPT21_active_list_201912*****")
filename = Application.GetOpenFilename(Title:="******Please choose your BR_AO*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Active List
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSES+DSHK - Others (DS-D Cust)")


'mySheetName.Range("E:E").EntireColumn.Insert
mySheetName.Range("B:B").EntireColumn.Insert

'With ThisWorkbook.mySheetName
 With mySheetName
            
    .Cells(1, 2) = "TRIM CIF No."
    .Cells(1, 2).Interior.ColorIndex = 6 'yellow
    .Cells(1, 2).WrapText = True
  
    
    For Each mycell In mySheetName.Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
        TmyCell = WorksheetFunction.Trim(mycell)
        
        .Cells(mycell.Row, "B") = TmyCell
        '.Cells(myCell.Row, "E").Interior.ColorIndex = 6 'yellow
    Next mycell
  
End With

'myFileName.Save

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in activesheet in Customer Enquiry_YYYYMMDD
'find last row in Column A in This Workbook ("PARTY_NUMBER.") in "DSDirectSheet1" in DSdirect Macro

'find last row in Column A ("PARTY_NUMBER")in BR_AO
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "B").End(xlUp).Row

'set the range for Vlookup all active rows and columns in BR_AO
Set myRangeName = mySheetName.Range("B1:B" & lastrow)


lastrow1 = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow1
    With currentSheet
        
        'If IsError(currentSheet.Cells(i, "H").Value) Then
        
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            
            .Cells(i, "U") = mySheetName.Cells(matchRow, "C").Value
            .Cells(i, "V") = mySheetName.Cells(matchRow, "D").Value
            .Cells(i, "W") = mySheetName.Cells(matchRow, "E").Value
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "F").Value
            '.Cells(i, "J") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "J").Value
            
            '.Cells(i, "M") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "L").Interior.ColorIndex = 7
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            
            'If .Cells(i, "G") <> "Closed" Then
            
            '.Cells(i, "R") = "Inactive(Unknown)"
            '.Cells(i, "M") = "N (Unknown)"
            '.Cells(i, "N") = "N (Unknown)"
            
            'Else
            .Cells(i, "U") = CVErr(xlErrNA)
            .Cells(i, "V") = CVErr(xlErrNA)
            .Cells(i, "W") = CVErr(xlErrNA)
            '.Cells(i, "P") = CVErr(xlErrNA)
            
            '.Cells(i, "L").Interior.ColorIndex = 8
            'End If
        End If
        
        'End If
    End With
Next i


myFileName.Close saveChanges:=False

End Sub


