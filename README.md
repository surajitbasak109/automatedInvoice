# automatedInvoice
Automated Professional Excel Invoice with VBA

## Instructions
If you are reading this file it means that you are interested to use this invoice system. So let's make some changes to make this invoice more efficient.

At first copy this directory to C:\ drive So it should be C:\Challan

So you need to rename this directory as "Challan" and move this directory to C(Primary) drive. After that create "Report" directory inside the Challan folder and now you are ready to use this automated Excdel invoice.

P.S. There may be some intentional bug in this, so, you are requested to issue that bug if arise. Or you can modify the VBA code so that it can be rectified.

## Code Snippets

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
