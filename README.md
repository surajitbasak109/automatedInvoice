# automatedInvoice
Automated Professional Excel Invoice with VBA

## Instructions
If you are reading this file it means that you are interested to use this invoice system. So let's make some changes to make this invoice more efficient.

At first copy this directory to C:\ drive So it should be C:\Challan

So you need to rename this directory as "Challan" and move this directory to C(Primary) drive. After that create "Report" directory inside the Challan folder and now you are ready to use this automated Excdel invoice.

P.S. There may be some intentional bug in this, so, you are requested to issue that bug if arise. Or you can modify the VBA code so that it can be rectified.

## Code Snippets

### Sheet1(Invoice)

```vba
Public myChng As String

Private Sub clearEntryBtn_Click()
    Call AreYouSure
End Sub

Private Sub ComboBox1_Change()
    Sheets("Invoice").Range("A7").Value = ComboBox1.Value
    Dim valueToEnter As String
    valueToEnter = getValue(ComboBox1.Value)
    Sheets("Invoice").Range("A8").Value = valueToEnter
End Sub
Function getValue(valueToFind) As Range
    Dim bottomCell As Range
    With Sheets("Customer Details")
        Set bottomCell = .Cells.Find(what:=valueToFind)
        Set getValue = bottomCell.Offset(0, 1)
        ' Now, your offsetCell has been created as a range, so go forth young padawan!
    End With
End Function

Private Sub SaveEntryBtn_Click()
    Call SaveEntry
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If myChng = "off" Then Exit Sub
    
    Dim amtInWrd As String
    amtInWrd = Module1.ConvertCurrencyToEnglish(Range("amtPayable"))
    
    If Not Application.Intersect(Range("itmQtyRgn"), Target) Is Nothing Then
        Range("totalQty") = Application.Sum(Range("itmQtyRgn"))
    End If

    If Not Application.Intersect(Range("itmPriceRgn"), Target) Is Nothing Then
        Target.Offset(0, 1) = Target.Value * Target.Offset(0, -2)
        Target.Offset(0, 2).Select
    End If
    
    If Not Application.Intersect(Range("itmDiscRgn"), Target) Is Nothing Then
        Target.Offset(0, 1) = Target.Offset(0, -1) * (1 - Target)
        Target.Offset(1, -5).Select
    End If
    
    If Not Application.Intersect(Range("itmGrsAmtRgn"), Target) Is Nothing Then
        Range("totalGrsAmt") = Application.Sum(Range("itmGrsAmtRgn"))
        Target.Offset(0, 2) = Target * (1 - Target.Offset(0, 1))
    End If
    
    If Not Application.Intersect(Range("itmNetAmtRgn"), Target) Is Nothing Then
        Range("totalNetAmt") = Application.Sum(Range("itmNetAmtRgn"))
    End If
    
    If Not Application.Intersect(Range("C11"), Target) Is Nothing Then
        Range("C43") = Range("C11")
    End If
    
    If Not Application.Intersect(Range("amtPayable"), Target) Is Nothing Then
        Range("amtInWords") = amtInWrd
    End If
    
    If Not Application.Intersect(Range("totalNetAmt"), Target) Is Nothing Then
        Range("roundOffAmt") = XLMod(Range("totalNetAmt"), 1)
    End If
    
    If Not Application.Intersect(Range("roundOffAmt"), Target) Is Nothing Then
        If Range("roundOffAmt") > "0.49" Then
            Range("billTotal") = Application.Ceiling(Range("totalNetAmt"), 1)
            Range("amtPayable") = Application.Ceiling(Range("totalNetAmt"), 1)
        End If
        If Range("roundOffAmt") < "0.49" Then
            Range("billTotal") = Application.Floor(Range("totalNetAmt"), 1)
            Range("amtPayable") = Application.Floor(Range("totalNetAmt"), 1)
        End If
    End If
    
    If Not Application.Intersect(Range("fareCharge"), Target) Is Nothing Then
        Range("amtPayable") = Range("billTotal") + Range("fareCharge")
    End If
    
    
End Sub

Sub clearEntry()
    myChng = "off"
    Range("billDate") = Date
    Range("customerDetails") = ""
    Range("billItm") = ""
    Range("billTrnAmt") = ""
    Range("billTrnAmt2") = ""
    Range("amtInWords") = ""
    myChng = "on"
End Sub

Sub AreYouSure()

    Dim Sure As Integer
    
    Sure = MsgBox("Are you sure? It will erase your bill entry.", vbOKCancel)
    If Sure = 1 Then Call clearEntry
End Sub

Function XLMod(a, b)
    ' This replicates the Excel MOD function
    XLMod = a - b * Int(a / b)
End Function

```
### ThisWorkBook
```vba
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call isEmpty
End Sub
```

### Module 1
```vba
Function ConvertCurrencyToEnglish(ByVal MyNumber)
' Edited by Karthikeyan karthikeyan@livetolearn.in
  Dim Temp
         Dim Rupees, Paise
         Dim DecimalPlace, count
 
         ReDim Place(9) As String
         Place(2) = " Thousand "
         Place(3) = " lakh "
         Place(4) = " Crore "
 
 
         ' Convert MyNumber to a string, trimming extra spaces.
         MyNumber = Trim(Str(MyNumber))
 
         ' Find decimal place.
         DecimalPlace = InStr(MyNumber, ".")
 
         ' If we find decimal place...
         If DecimalPlace > 0 Then
            ' Convert Paise
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            ' Hi! Note the above line Mid function it gives right portion
            ' after the decimal point
            'if only . and no numbers such as 789. accures, mid returns nothing
            ' to avoid error we added 00
            ' Left function gives only left portion of the string with specified places here 2
 
 
            Paise = ConvertTens(Temp)
 
 
            ' Strip off paise from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
         End If
 
         count = 1
        If MyNumber <> "" Then
 
            ' Convert last 3 digits of MyNumber to Indian Rupees.
            Temp = ConvertHundreds(Right(MyNumber, 3))
 
            If Temp <> "" Then Rupees = Temp & Place(count) & Rupees
 
            If Len(MyNumber) > 3 Then
               ' Remove last 3 converted digits from MyNumber.
               MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
               MyNumber = ""
            End If
 
        End If
 
            ' convert last two digits to of mynumber
            count = 2
 
            Do While MyNumber <> ""
            Temp = ConvertTens(Right("0" & MyNumber, 2))
 
            If Temp <> "" Then Rupees = Temp & Place(count) & Rupees
            If Len(MyNumber) > 2 Then
               ' Remove last 2 converted digits from MyNumber.
               MyNumber = Left(MyNumber, Len(MyNumber) - 2)
 
            Else
               MyNumber = ""
            End If
            count = count + 1
 
            Loop
 
 
 
 
         ' Clean up rupees.
         Select Case Rupees
            Case ""
               Rupees = ""
            Case "One"
               Rupees = "Rupee One"
            Case Else
               Rupees = "Rupees " & Rupees
         End Select
 
         ' Clean up paise.
         Select Case Paise
            Case ""
               Paise = ""
            Case "One"
               Paise = "One Paise"
            Case Else
               Paise = Paise & " Paise"
         End Select
 
         If Rupees = "" Then
         ConvertCurrencyToEnglish = Paise & " Only"
         ElseIf Paise = "" Then
         ConvertCurrencyToEnglish = Rupees & " Only"
         Else
         ConvertCurrencyToEnglish = Rupees & " and " & Paise & " Only"
         End If
 
End Function
 
 
Private Function ConvertDigit(ByVal MyDigit)
        Select Case Val(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select
 
End Function
 
Private Function ConvertHundreds(ByVal MyNumber)
    Dim Result As String
 
         ' Exit if there is nothing to convert.
         If Val(MyNumber) = 0 Then Exit Function
 
         ' Append leading zeros to number.
         MyNumber = Right("000" & MyNumber, 3)
 
         ' Do we have a hundreds place digit to convert?
         If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
         End If
 
         ' Do we have a tens place digit to convert?
         If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
         Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
         End If
 
         ConvertHundreds = Trim(Result)
End Function
 
 
Private Function ConvertTens(ByVal MyTens)
          Dim Result As String
 
         ' Is value between 10 and 19?
         If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
               Case 10: Result = "Ten"
               Case 11: Result = "Eleven"
               Case 12: Result = "Twelve"
               Case 13: Result = "Thirteen"
               Case 14: Result = "Fourteen"
               Case 15: Result = "Fifteen"
               Case 16: Result = "Sixteen"
               Case 17: Result = "Seventeen"
               Case 18: Result = "Eighteen"
               Case 19: Result = "Nineteen"
               Case Else
            End Select
         Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select
 
            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If
 
         ConvertTens = Result
End Function

```

### Module 2
```vba
Sub SaveEntry()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = Worksheets("Invoice")
    Set ws2 = Worksheets("Details")
    Dim erow As Long
    Dim count As Long
    Dim i As Integer
    Dim r As Long
    count = 0
    
    Application.DisplayAlerts = False
    
    If ws1.Range("customerName") = "" Then
        MsgBox ("Customer Name cannot be empty.")
        Exit Sub
    End If
    
    For i = 11 To 40
        If ws1.Cells(i, 2) = "" Then
            count = count + 1
        End If
    
    Next i
        If count = 30 Then
            MsgBox ("You cannot save empty bill")
            Exit Sub
        End If
    erow = ws2.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
    
    ws2.Activate
    Cells(erow, 1).Value = ws1.Range("customerName")
    Cells(erow, 2).Value = ws1.Range("billNum")
    Cells(erow, 3).Value = Format(ws1.Range("billDate"), "dd-mmm-yy")
    Cells(erow, 4).Value = ws1.Range("billType")
    Cells(erow, 5).Value = ws1.Range("amtPayable")
    
    ws1.Activate
    Dim Path As String
    Path = "C:\Challan\Reports\"
    ActiveWorkbook.Save
    ActiveWorkbook.ActiveSheet.SaveAs Filename:=Path & Replace(Range("customerName"), " ", "-") & "-" & Range("billNum") & ".xlsx", FileFormat:=51
    ActiveWorkbook.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Path & Replace(Range("customerName"), " ", "-") & "-" & Range("billNum") & ".pdf", OpenAfterPublish:=False
    Workbooks.Open Filename:="C:\Challan\Bill_Challan.xlsm"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub isEmpty()
    Dim ws1 As Worksheet
    Set ws1 = Worksheets("Invoice")
    
    Application.DisplayAlerts = False
    
    If WorksheetFunction.CountA(ws1.Range("invoiceEntry")) = 0 Then
        ActiveWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Exit Sub
    Else
        ActiveWorkbook.Close SaveChanges:=True
        Application.DisplayAlerts = False
    End If
End Sub

Sub PasteSpecial()
'
' PasteSpecial Macro
' Press Ctrl + Shift + V to paste unformatted text or values.
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
```
