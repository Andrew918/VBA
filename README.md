Sub ReferralBonus()

Application.DisplayAlerts = False

ThisWorkbook.Sheets("Step 1").Copy
ActiveWorkbook.SaveAs Filename:="C:\Users\Public\Desktop\Submissions_02.08.19.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=True
ActiveWorkbook.Close

ThisWorkbook.Sheets("Step 1").Copy
ActiveWorkbook.SaveAs Filename:="C:\Users\Public\Desktop\EE Referral Bonus Pymt02.08.19 PAYCK.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=True
ActiveWorkbook.Close

Application.DisplayAlerts = True                            'Create the 1st and 2nd uploading files'



Dim LastRow As Integer

Set dsheet = ThisWorkbook.Sheets("Step 1")

LastRow = dsheet.Cells(Rows.count, 1).End(xlUp).Row

For i = LastRow To 2 Step -1

If (dsheet.Cells(i, 6) = "Inactive") _
   Or (dsheet.Cells(i, 16) = "Inactive") _
   Or (dsheet.Cells(i, 3) = "Rising Star" And dsheet.Cells(i, 5) = "IC") _
   Or (dsheet.Cells(i, 3) = "Rising Star" And dsheet.Cells(i, 5).Value Like "IS*") _
   Or (dsheet.Cells(i, 3).Value = "" And dsheet.Cells(i, 5) = "IC") _
   Or (dsheet.Cells(i, 3).Value = "" And dsheet.Cells(i, 5).Value Like "IS*") _
   Or (dsheet.Cells(i, 13) = "Rising Star" And dsheet.Cells(i, 15) = "IC") _
   Or (dsheet.Cells(i, 13) = "Rising Star" And dsheet.Cells(i, 15).Value Like "IS*") _
   Or (dsheet.Cells(i, 13).Value = "" And dsheet.Cells(i, 15) = "IC") _
   Or (dsheet.Cells(i, 13).Value = "" And dsheet.Cells(i, 15).Value Like "IS*") _
   Or (Not dsheet.Cells(i, 25).Value = "" And Not dsheet.Cells(i, 27).Value = "") _
   Or (dsheet.Cells(i, 15) = "IC" And dsheet.Cells(i, 19).Value < 30) _
   Or (dsheet.Cells(i, 15) = "IC" And dsheet.Cells(i, 19).Value >= 30 And dsheet.Cells(i, 19).Value < 60 And Not dsheet.Cells(i, 25).Value = "") _
   Or (dsheet.Cells(i, 15).Value Like "IS*" And dsheet.Cells(i, 19).Value < 30) _
   Or (dsheet.Cells(i, 15) Like "IS*" And dsheet.Cells(i, 19).Value >= 30 And dsheet.Cells(i, 19).Value < 60 And Not dsheet.Cells(i, 25).Value = "") _
   Or (Not dsheet.Cells(i, 15).Value Like "IS*" And Not dsheet.Cells(i, 15).Value = "IC" And dsheet.Cells(i, 19).Value < 90) _
   Or (Not dsheet.Cells(i, 15).Value Like "IS*" And Not dsheet.Cells(i, 15).Value = "IC" And dsheet.Cells(i, 19).Value >= 90 And dsheet.Cells(i, 19).Value < 180 And Not dsheet.Cells(i, 25).Value = "") _
Then
   dsheet.Rows(i).EntireRow.delete

End If

Next i                                                       'Delete rows'



For x = 2 To LastRow

If (dsheet.Cells(x, 15).Value Like "IS*" Or dsheet.Cells(x, 15).Value = "IC") _
   And dsheet.Cells(x, 19).Value >= 30 _
   And dsheet.Cells(x, 19).Value < 60 _
   And dsheet.Cells(x, 25).Value = "" _
Then
   dsheet.Cells(x, 25) = 100
   dsheet.Cells(x, 25).Font.Color = vbRed
   dsheet.Cells(x, 26) = Date + 6
   dsheet.Cells(x, 26).Font.Color = vbRed

ElseIf (dsheet.Cells(x, 15).Value Like "IS*" Or dsheet.Cells(x, 15).Value = "IC") _
   And dsheet.Cells(x, 19).Value >= 60 _
   And dsheet.Cells(x, 25).Value = "" _
Then
   dsheet.Cells(x, 25) = 100
   dsheet.Cells(x, 25).Font.Color = vbRed
   dsheet.Cells(x, 26) = Date + 6
   dsheet.Cells(x, 26).Font.Color = vbRed
   dsheet.Cells(x, 27) = 100
   dsheet.Cells(x, 27).Font.Color = vbRed
   dsheet.Cells(x, 28) = Date + 6
   dsheet.Cells(x, 28).Font.Color = vbRed
   
ElseIf (dsheet.Cells(x, 15).Value Like "IS*" Or dsheet.Cells(x, 15).Value = "IC") _
   And dsheet.Cells(x, 19).Value >= 60 _
   And Not dsheet.Cells(x, 25).Value = "" _
Then
   dsheet.Cells(x, 27) = 100
   dsheet.Cells(x, 27).Font.Color = vbRed
   dsheet.Cells(x, 28) = Date + 6
   dsheet.Cells(x, 28).Font.Color = vbRed

End If

Next x                                                     'Summarize the ICs and ISs'


For y = 2 To LastRow

If Not dsheet.Cells(y, 15).Value Like "IS*" _
   And Not dsheet.Cells(y, 15).Value = "IC" _
   And dsheet.Cells(y, 19).Value >= 90 _
   And dsheet.Cells(y, 19).Value < 180 _
   And dsheet.Cells(y, 25).Value = "" _
Then
   dsheet.Cells(y, 25) = 250
   dsheet.Cells(y, 25).Font.Color = vbRed
   dsheet.Cells(y, 26) = Date + 6
   dsheet.Cells(y, 26).Font.Color = vbRed

ElseIf Not dsheet.Cells(y, 15).Value Like "IS*" _
   And Not dsheet.Cells(y, 15).Value = "IC" _
   And dsheet.Cells(y, 19).Value >= 180 _
   And dsheet.Cells(y, 25).Value = "" _
Then
   dsheet.Cells(y, 25) = 250
   dsheet.Cells(y, 25).Font.Color = vbRed
   dsheet.Cells(y, 26) = Date + 6
   dsheet.Cells(y, 26).Font.Color = vbRed
   dsheet.Cells(y, 27) = 250
   dsheet.Cells(y, 27).Font.Color = vbRed
   dsheet.Cells(y, 28) = Date + 6
   dsheet.Cells(y, 28).Font.Color = vbRed
   
ElseIf Not dsheet.Cells(y, 15).Value Like "IS*" _
   And Not dsheet.Cells(y, 15).Value = "IC" _
   And dsheet.Cells(y, 19).Value >= 180 _
   And Not dsheet.Cells(y, 25).Value = "" _
Then
   dsheet.Cells(y, 27) = 250
   dsheet.Cells(y, 27).Font.Color = vbRed
   dsheet.Cells(y, 28) = Date + 6
   dsheet.Cells(y, 28).Font.Color = vbRed

End If

Next y                                                    'Summarize the salaried employees'



ThisWorkbook.Sheets.Add After:=ActiveSheet
ThisWorkbook.Sheets("Sheet1").Name = "SendtoPayroll"
ThisWorkbook.Sheets.Add After:=ActiveSheet
ThisWorkbook.Sheets("Sheet2").Name = "SendtoDoan"
    

ThisWorkbook.Sheets("Step 1").Range("A:B").Copy Destination:=Sheets("SendtoPayroll").Range("A:B")
ThisWorkbook.Sheets("Step 1").Range("K:L").Copy Destination:=Sheets("SendtoPayroll").Range("C:D")
ThisWorkbook.Sheets("Step 1").Range("Y:AB").Copy Destination:=Sheets("SendtoPayroll").Range("E:H")

ThisWorkbook.Sheets("Step 1").Range("A:B").Copy Destination:=Sheets("SendtoDoan").Range("A:B")
ThisWorkbook.Sheets("Step 1").Range("K:L").Copy Destination:=Sheets("SendtoDoan").Range("C:D")
ThisWorkbook.Sheets("Step 1").Range("Y:AB").Copy Destination:=Sheets("SendtoDoan").Range("E:H")


Set csheet = ThisWorkbook.Sheets("SendtoPayroll")

csheet.Cells(1, 9) = "Payment Amount"
csheet.Cells(1, 1).Font.Color = vbWhite
csheet.Cells(1, 9).Font.Color = vbWhite
csheet.Cells(1, 1).Interior.ColorIndex = 1
csheet.Cells(1, 9).Interior.ColorIndex = 1

LastRow = csheet.Cells(Rows.count, 1).End(xlUp).Row

For i = 2 To LastRow

If csheet.Cells(i, 5).Font.Color = vbRed _
Then
   csheet.Cells(i, 9) = csheet.Cells(i, 5) + csheet.Cells(i, 7)

ElseIf csheet.Cells(i, 5).Font.Color = vbBlack _
Then
   csheet.Cells(i, 9) = csheet.Cells(i, 7)

End If

Next i


csheet.Columns("B:H").EntireColumn.delete


For m = LastRow To 2 Step -1
    For n = 2 To LastRow
        If csheet.Cells(m, 1).Value = csheet.Cells(n, 1).Value _
        And m > n _
        Then
            csheet.Cells(n, 2).Value = csheet.Cells(m, 2).Value + csheet.Cells(n, 2).Value
            csheet.Rows(m).EntireRow.delete
            Exit For
        End If
    Next n
Next m

                                                               'Consolidate Columns'

                                                               
Application.DisplayAlerts = False

ThisWorkbook.Sheets("SendtoPayroll").Copy
ActiveWorkbook.SaveAs Filename:="C:\Users\Public\Desktop\EE Referral Bonus Pymt 02.08.19 PAYCK-PAYROLL.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=True
ActiveWorkbook.Close

Application.DisplayAlerts = True                               '3rd report send to Payroll'




Application.DisplayAlerts = False

ThisWorkbook.Sheets("SendtoDoan").Copy
ActiveWorkbook.SaveAs Filename:="C:\Users\Public\Desktop\UPLOAD EEReferralPaymentData02.08.19.csv", FileFormat:=xlCSV, CreateBackup:=True
ActiveWorkbook.Close

Application.DisplayAlerts = True                               'Final report send to Doan'


Workbooks.Open "C:\Users\Public\Desktop\EE Referral Bonus Pymt 02.08.19 PAYCK-PAYROLL.xlsx"
MsgBox ("Good job! And do we get anyone today?")

End Sub


------------------------
Sub consolidate()

Dim Sh As Worksheet
    Dim LastRow As Long
    Dim Rng As Range
    Set Sh = ThisWorkbook.Worksheets("Sendtopayroll")
    Sh.Columns(5).Insert
    LastRow = Sh.Range("A65536").End(xlUp).Row
    With Sh.Range("A1:A" & LastRow).Offset(0, 4)
        .FormulaR1C1 = "=IF(COUNTIF(R1C[-4]:RC[-4],RC[-4])>1,"""",SUMIF(R1C[-4]:R[" & LastRow & "]C[-4],RC[-4],R1C[-1]:R[" & LastRow & "]C[-1]))"
        .Value = .Value
    End With
    Sh.Columns(4).delete
    Sh.Rows(1).Insert
    Set Rng = Sh.Range("D1:D" & LastRow + 1)
    With Rng
        .AutoFilter Field:=1, Criteria1:="="
        .SpecialCells(xlCellTypeVisible).EntireRow.delete
    End With

End Sub

---------------------------------------------

Sub Check()
'
' Check Macro
'

'
    Sheets("FeedExecSum").Select
    ActiveWorkbook.SlicerCaches("Slicer_Campaign2").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_CAMPAIGN3").ClearManualFilter
    Sheets("Executive Summary 2020").Select
    Range("A4").Select
End Sub
Sub Uncheck()
'
' Uncheck Macro
'

'
    Sheets("FeedExecSum").Select
    ActiveWorkbook.SlicerCaches("Slicer_Campaign2").VisibleSlicerItemsList = Array _
        ( _
        "[Table2].[Campaign].&[Cancel B4 Install]")
    ActiveWorkbook.SlicerCaches("Slicer_CAMPAIGN3").VisibleSlicerItemsList = Array _
        ( _
        "[Table1].[CAMPAIGN].&[Saved Cart]", "[Table1].[CAMPAIGN].&[SISFA Call Center]" _
        , "[Table1].[CAMPAIGN].&[SISFA Store Visit]")
    Sheets("Executive Summary 2020").Select
    Range("A4").Select
End Sub

-----------------------------------------------------------
Sub UnhideSheetsforReportUpdate()
'
' UnhideSheetsforReportUpdate Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'

    Sheets("EM DATA").Visible = True
    Sheets("DM DATA").Visible = True
    Sheets("DM Tables").Visible = True
    Sheets("EM Tables").Visible = True
    Sheets("VLOOKUPs").Visible = True
End Sub
---------------------------------------------------------------
