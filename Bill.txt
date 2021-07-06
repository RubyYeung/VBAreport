Sub Step1clearcontents()



'With ThisWorkbook.Worksheets("ServiceFeesheet1").ShowAllData
If Worksheets("ServiceFeesheet1").FilterMode = True Then
            Worksheets("ServiceFeesheet1").ShowAllData
End If

With ThisWorkbook.Worksheets("ServiceFeesheet1").UsedRange
    'With .Cells(1, 1).CurrentRegion
        'With .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)
            .Cells.clearcontents
            .Cells.Interior.Pattern = xlNone
            '.Cells.Font.ColorIndex = xlAutomatic
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
            '.Cells(1, 12) = "Customer Type Original"
            '.Cells(1, 13) = "Customer Type 1st update"
            '.Cells(1, 14) = "Customer Type Top20 Only"
            '.Cells(1, 15) = "Customer Type Final"
            '.Cells(1, 16) = "% of Total CASA"
            '.Cells(1, 17) = "% of Total CASA (Without Personal)"
            '.Cells(1, 18) = "Bills - Amount Range"
            '.Cells(1, 19) = "Non-Bills Borrowing - Amount Range"
            '.Cells(1, 20) = "Non-Borrowing - Amount Range"
            '.Cells(1, 21) = "Tier Interest Client"
        'End With
    'End With
End With

End Sub



Sub Step2importDSDirectServiceFee()

Dim wb1 As Workbook
Dim wb2 As Workbook
Dim Sheet As Worksheet
Dim PasteStart As Range

Set wb1 = ActiveWorkbook
Set PasteStart = [ServiceFeeSheet1!A1]

FileToOpen = Application.GetOpenFilename _
(Title:="******Please choose THIS MONTH Service Fee Transaction Maintenance_YYYYMM-YYYYMM*****", _
FileFilter:="Report Files *.csv(*.csv),")

If FileToOpen = False Then
    MsgBox "No File Specified.", vbExclamation, "ERROR"
    Exit Sub
Else
    Set wb2 = Workbooks.Open(filename:=FileToOpen)


    With ThisWorkbook.Worksheets("ServiceFeeSheet1")

        .Cells(1, 2) = "Company Name"
        .Cells(1, 3) = "BR #"
        .Cells(1, 4) = "Voucher Description"
        .Cells(1, 6) = "Statement Desc."
        .Cells(1, 14) = "RH"
        .Cells(1, 15) = "OFFICER_CD"
        .Cells(1, 16) = "OFFICER_SUB_CD"
        .Cells(1, 17) = "Follow Up Action"
        .Cells(1, 18) = "Pending IM's Review & Approval (per OSD's Action)"
        .Cells(1, 19) = "OSD's Action / Feedback"
        .Columns("P").NumberFormat = "General"
        
    End With

    For Each Sheet In wb2.Sheets
   
        With Range("A1", Range("A" & Rows.Count).End(xlUp))
            '.AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
            'On Error Resume Next
            '.Offset(0, 0).EntireColumn.Copy PasteStart
            
        
            
            .Offset(0, 0).EntireColumn.Copy PasteStart.Range("A1")
            .Offset(0, 1).EntireColumn.Copy PasteStart.Range("E1")
            .Offset(0, 2).EntireColumn.Copy PasteStart.Range("G1")
            .Offset(0, 3).EntireColumn.Copy PasteStart.Range("H1")
            .Offset(0, 4).EntireColumn.Copy PasteStart.Range("I1")
            .Offset(0, 5).EntireColumn.Copy PasteStart.Range("J1")
            .Offset(0, 6).EntireColumn.Copy PasteStart.Range("K1")
            .Offset(0, 7).EntireColumn.Copy PasteStart.Range("L1")
            .Offset(0, 8).EntireColumn.Copy PasteStart.Range("M1")
        End With
    Next Sheet
    
    
    With ThisWorkbook.Worksheets("ServiceFeeSheet1")

        .Range("A1:M1").Interior.ColorIndex = 6 'yellow
        .Range("N1:P1").Interior.ColorIndex = 45 'orange
        .Range("Q1:Q1").Interior.ColorIndex = 40 'pale orange
        .Range("R1:R1").Interior.ColorIndex = 34 'pale blue
        .Range("S1:S1").Interior.ColorIndex = 35 'pale green
                
        .Cells(1, 4).Font.ColorIndex = 45 'orange
        .Cells(1, 6).Font.ColorIndex = 45 'orange
        .Cells(1, 8).Font.ColorIndex = 45 'orange
        .Cells(1, 9).Font.ColorIndex = 45 'orange
        .Cells(1, 10).Font.ColorIndex = 45 'orange
        .Cells(1, 17).Font.ColorIndex = 45 'orange
        '.Cells(1, 17) = "Follow Up Action"
        '.Cells(1, 18) = "Pending IM's Review & Approval (per OSD's Action)"
        '.Cells(1, 19) = "OSD's Action / Feedback"
        
        .Columns("J").NumberFormat = "general"
        .Columns("J").Value = .Columns("J").Value
    End With
    
End If

    wb2.Close Save = False
    
End Sub


Sub Step3DSDirectServiceFeevlookup()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("ServiceFeeSheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose Customer Enquiry_YYYYMMDD FINAL*****")
filename = Application.GetOpenFilename(Title:="******Please choose THIS MONTH Customer Enquiry_YYYYMMDD FINAL*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in Customer Enquiry_YYYYMMDD FINAL
'Set mySheetName = myFileName.ActiveSheet
Set mySheetName = myFileName.Worksheets("DSDirectSheet1")

'find last row in Column A ("GROUP ID")in Customer Enquiry_YYYYMMDD FINAL
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in Customer Enquiry_YYYYMMDD FINAL
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("GROUP ID") in "ServiceFeeSheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            '.Cells(i, "B") = mySheetName.Cells(matchRow, "C").Value
            .Cells(i, "C") = mySheetName.Cells(matchRow, "D").Value
            .Cells(i, "D") = Replace(mySheetName.Cells(matchRow, "D").Value, " ", "")
                                   
            .Cells(i, "D").Font.ColorIndex = 45 'orange
            
            .Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            .Cells(i, "O") = mySheetName.Cells(matchRow, "J").Value
            .Cells(i, "P") = mySheetName.Cells(matchRow, "K").Value
            .Cells(i, "P").NumberFormat = "General"
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            '.Cells(i, "B") = CVErr(xlErrNA)
            .Cells(i, "C") = CVErr(xlErrNA)
            .Cells(i, "D") = CVErr(xlErrNA)
            .Cells(i, "N") = CVErr(xlErrNA)
            .Cells(i, "O") = CVErr(xlErrNA)
            .Cells(i, "P") = CVErr(xlErrNA)
        End If

    End With
Next i


myFileName.Close Save = False

End Sub

Sub Step4ITfilelookup()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("ServiceFeeSheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose THIS MONTH file provided by IT Simon*****")
filename = Application.GetOpenFilename(Title:="******Please choose THIS MONTH file provided by IT Simon*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in file provided by IT Simon
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("DSDirectSheet1")

'find last row in Column A ("GROUP ID")in file provided by IT Simon
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "A").End(xlUp).Row

'set the range for Vlookup all active rows and columns in file provided by IT Simon
Set myRangeName = mySheetName.Range("A1:A" & lastrow)

' find last row in Column A in This Workbook ("GROUP ID") in "ServiceFeeSheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "A"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "A"), myRangeName, 0)
            .Cells(i, "B") = mySheetName.Cells(matchRow, "C").Value
            '.Cells(i, "C") = mySheetName.Cells(matchRow, "D").Value
            '.Cells(i, "D") = mySheetName.Cells(matchRow, "D").Value
            '.Cells(i, "D").Font.ColorIndex = 45 'orange
            
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "O") = mySheetName.Cells(matchRow, "J").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "K").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            .Cells(i, "B") = CVErr(xlErrNA)
            '.Cells(i, "C") = CVErr(xlErrNA)
            '.Cells(i, "D") = CVErr(xlErrNA)
            '.Cells(i, "N") = CVErr(xlErrNA)
            '.Cells(i, "O") = CVErr(xlErrNA)
            '.Cells(i, "P") = CVErr(xlErrNA)
        End If

    End With
Next i


myFileName.Close Save = False

End Sub

Sub Step5AddStatementDsec()


Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet
Set currentSheet = ThisWorkbook.Worksheets("ServiceFeeSheet1")


' find last row in Column A in This Workbook ("GROUP ID") in "ServiceFeeSheet1"
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If .Cells(i, "G") = "Monthly Fee" Then
           .Cells(i, "F") = "DS-DIRECT MONTHLY FEE"
           .Cells(i, "F").Font.ColorIndex = 45 'orange
        ElseIf .Cells(i, "G") = "Setup Fee" Then
           .Cells(i, "F") = "DS-DIRECT SETUP FEE"
           .Cells(i, "F").Font.ColorIndex = 45 'orange
        Else
            .Cells(i, "F") = CVErr(xlErrNA)
            .Cells(i, "F").Font.ColorIndex = 45 'orange
            '.Cells(i, "D") = mySheetName.Cells(matchRow, "D").Value
            '.Cells(i, "N") = mySheetName.Cells(matchRow, "I").Value
            '.Cells(i, "O") = mySheetName.Cells(matchRow, "J").Value
            '.Cells(i, "P") = mySheetName.Cells(matchRow, "K").Value
        'Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            '.Cells(i, "B") = CVErr(xlErrNA)
            '.Cells(i, "C") = CVErr(xlErrNA)
            '.Cells(i, "D") = CVErr(xlErrNA)
            '.Cells(i, "N") = CVErr(xlErrNA)
            '.Cells(i, "O") = CVErr(xlErrNA)
            '.Cells(i, "P") = CVErr(xlErrNA)
        End If

    End With
Next i

End Sub
Sub Step6DSdirectBillingVlookupfilefromReplyFromOSD()

Dim filename                As String
Dim myFileName              As Workbook
Dim currentSheet            As Worksheet
Dim mySheetName             As Worksheet
Dim myRangeName             As Range
Dim lastrow                 As Long
Dim i                       As Long
Dim matchRow                As Long


'set current worksheet of activesheet in Service Fee Transaction Maintenance_YYYYMM-YYYYMM for IM
'Set currentSheet = ThisWorkbook.Worksheets("CASASheet1")
Set currentSheet = ThisWorkbook.Worksheets("ServiceFeeSheet1")

'get workbook path
'filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="******Please choose your Current Month Service Fee Transaction Maintenance_YYYYMM-YYYYMM_reply from OSD*******")
filename = Application.GetOpenFilename(Title:="******Please choose your Current Month Service Fee Transaction Maintenance_YYYYMM-YYYYMM_reply from OSD*****")


'set our workbook and open it
Set myFileName = Application.Workbooks.Open(filename)

'set searched worksheet in "Service Fee Transaction Maintenance_YYYYMM-YYYYMM_for OSD_reply from OSD"
Set mySheetName = myFileName.ActiveSheet
'Set mySheetName = myFileName.Worksheets("CASASheet1")

'find last row in Column A ("PARTY_NUMBER")in "Service Fee Transaction Maintenance_YYYYMM-YYYYMM_for OSD_reply from OSD"
lastrow = mySheetName.Cells(mySheetName.Rows.Count, "F").End(xlUp).Row

'set the range for Vlookup all active rows and columns in "Service Fee Transaction Maintenance_YYYYMM-YYYYMM_for OSD_reply from OSD"
Set myRangeName = mySheetName.Range("F1:F" & lastrow)

'Set myRangeName1 = mySheetName.Range("M1:M" & lastrow)

' find last row in Column A in This Workbook ("PARTY_NUMBER.") in Service Fee Transaction Maintenance_YYYYMM-YYYYMM for IM
lastrow = currentSheet.Cells(currentSheet.Rows.Count, "E").End(xlUp).Row
'MsgBox (lastRow)

For i = 2 To lastrow
    With currentSheet
        If Not IsError(Application.Match(.Cells(i, "E"), myRangeName, 0)) Then
            matchRow = Application.Match(.Cells(i, "E"), myRangeName, 0)
            .Cells(i, "Q") = mySheetName.Cells(matchRow, "R").Value
            .Cells(i, "R") = mySheetName.Cells(matchRow, "S").Value
            .Cells(i, "S") = mySheetName.Cells(matchRow, "T").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "K").Value
            '.Cells(i, "K") = mySheetName.Cells(matchRow, "G").Value
        Else ' Item No. record not found
            ' put #NA in cells, to know it's not found
            ''''.Cells(i, "Q") = CVErr(xlErrNA)
            ''''.Cells(i, "R") = CVErr(xlErrNA)
            ''''.Cells(i, "S") = CVErr(xlErrNA)
            '.Cells(i, "K") = CVErr(xlErrNA)
            '.Cells(i, "K") = CVErr(xlErrNA)
        End If

    End With
Next i

myFileName.Close Save = False

End Sub

