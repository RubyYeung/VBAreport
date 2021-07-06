Sub SortingOFFICER_CD()
Dim PvtTbl As PivotTable
Set PvtTbl = Worksheets("Pivot").PivotTables("PivotTable4")

'On Error Resume Next
ThisWorkbook.Worksheets("Data Validate").Select
For i = 2 To ThisWorkbook.Worksheets("Data Validate").Range("B1", Range("B1").End(xlDown)).Rows.Count

PvtTbl.PivotFields("Officer_CD").PivotItems(Worksheets("Data Validate").Cells(i, 2).Value).Position = i - 1
On Error Resume Next
'PvtTbl.PivotFields("Officer_CD").PivotItems("3910").Position = 1
'PvtTbl.PivotFields("Officer_CD").PivotItems("3912").Position = 2
'PvtTbl.PivotFields("Officer_CD").PivotItems("3916").Position = 3
'PvtTbl.PivotFields("Officer_CD").PivotItems("3986").Position = 4
'PvtTbl.PivotFields("Officer_CD").PivotItems("3988").Position = 5
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B6").Position = 6
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BD").Position = 7
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BN").Position = 8
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BT").Position = 9

'PvtTbl.PivotFields("Officer_CD").PivotItems("3911").Position = 10
'PvtTbl.PivotFields("Officer_CD").PivotItems("3987").Position = 11
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B9").Position = 12

'PvtTbl.PivotFields("Officer_CD").PivotItems("3936").Position = 13
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B1").Position = 14
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BL").Position = 15
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BM").Position = 16
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BQ").Position = 17

'PvtTbl.PivotFields("Officer_CD").PivotItems("39B2").Position = 18
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BA").Position = 19
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BP").Position = 20
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BV").Position = 21
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BW").Position = 22


'PvtTbl.PivotFields("Officer_CD").PivotItems("3938").Position = 23
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B0").Position = 24
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B7").Position = 25
'PvtTbl.PivotFields("Officer_CD").PivotItems("39B8").Position = 26
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BE").Position = 27
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BF").Position = 28
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BS").Position = 29

'PvtTbl.PivotFields("Officer_CD").PivotItems("3937").Position = 30
'PvtTbl.PivotFields("Officer_CD").PivotItems("3983").Position = 31
'''''PvtTbl.PivotFields("Officer_CD").PivotItems("39BJ").Position = 32
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BG").Position = 33
'PvtTbl.PivotFields("Officer_CD").PivotItems("39BR").Position = 34

'PvtTbl.PivotFields("Officer_CD").PivotItems("3909").Position = 35

''''''PvtTbl.PivotFields("Officer_CD").PivotItems("3917").Position = 36
'PvtTbl.PivotFields("Officer_CD").PivotItems("3918").Position = 37
'PvtTbl.PivotFields("Officer_CD").PivotItems("3919").Position = 38

'PvtTbl.PivotFields("Officer_CD").PivotItems("ZZZZ").Position = 39

Next
End Sub

Sub SortingOFFICER_TEAM_HEAD_NAME()
Dim PvtTbl As PivotTable
Set PvtTbl = Worksheets("Pivot").PivotTables("PivotTable4")

ThisWorkbook.Worksheets("Data Validate").Select
For i = 2 To Worksheets("Data Validate").Range("A1", Range("A1").End(xlDown)).Rows.Count

PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems(Worksheets("Data Validate").Cells(i, 1).Value).Position = i - 1
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("DAVID LOOK").Position = 1
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("JOHN LAW").Position = 2
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("DESMOND LAM").Position = 3
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("LEO KWOK").Position = 4
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("JOE CHAN").Position = 5
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("BERNARD LEE").Position = 6
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("OTHERS").Position = 7
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("SAM").Position = 8
'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("SYBIL CHAN").Position = 9

'PvtTbl.PivotFields("OFFICER_REGION_HEAD_NAME").PivotItems("CBD OTHERS").Position = 10

Next
End Sub

