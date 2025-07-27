Attribute VB_Name = "Module1"
Public wshWorkSheet As Worksheet

'theVariable(y, x)
'y is for Total, PrePay, PostPay
'x is for Cost, VAT, Cost+VAT
Public dMoney(2, 2) As Double
Public iBillions(2, 2) As Integer
Public iMillions(2, 2) As Integer
Public iThousands(2, 2) As Integer
Public iRoubles(2, 2) As Integer
Public iKopecks(2, 2) As Integer

Public dTaxRate As Double
Public dPayRate As Double

Private Sub ShowField()
Dim btnCalculate As Button, btnClean As Button
Dim btnCopy(2) As Button
Dim r As Range
Application.ScreenUpdating = False

Set wshWorkSheet = ActiveSheet

With wshWorkSheet.Cells(3, 1).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
    Formula1:="����� ���� ���, � ��� ����� ���" '�� ���������� ��������� ������� ����� "����� ����"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
End With

wshWorkSheet.Buttons.Delete

Set r = wshWorkSheet.Cells(7, 6)
Set btnCalculate = wshWorkSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
With btnCalculate
    .OnAction = "Calculate"
    .Caption = "����������"
    .Name = "Calculate"
End With
    
Set r = wshWorkSheet.Cells(9, 6)
Set btnClean = wshWorkSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
With btnClean
    .OnAction = "ClearAll"
    .Caption = "��������"
    .Name = "ClearAll"
End With

Set r = wshWorkSheet.Cells(14, 6)
Set btnClean = wshWorkSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
With btnClean
    .OnAction = "CopyTotal"
    .Caption = "����������"
    .Name = "CopyTotal"
End With

Set r = wshWorkSheet.Cells(21, 6)
Set btnClean = wshWorkSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
With btnClean
    .OnAction = "CopyPrePay"
    .Caption = "����������"
    .Name = "CopyPrePay"
End With

Set r = wshWorkSheet.Cells(28, 6)
Set btnClean = wshWorkSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height)
With btnClean
    .OnAction = "CopyPostPay"
    .Caption = "����������"
    .Name = "CopyPostPay"
End With

Application.ScreenUpdating = True
End Sub

Private Sub Calculate()
Set wshWorkSheet = ActiveSheet
If wshWorkSheet.Cells(3, 1).Value = "����� ���� ���" Then '�� ���������� ��������� ������� ����� "����� ����"
    Call AddVAT
ElseIf wshWorkSheet.Cells(3, 1).Value = "� ��� ����� ���" Then
    Call SubVAT
Else
    MsgBox ("��� ������ ���� ��� �� ��� �����?")
End If
End Sub

Private Sub AddVAT()
If wshWorkSheet.Cells(2, 2).Value > 833333333333.32 Then
    MsgBox ("��������� ������ ���� ������" & Chr(13) & "833 ����. 333 ���. 333 ���. 333 ���. 33 ���.")
    Exit Sub
End If
If wshWorkSheet.Cells(2, 2).Value < 0.01 Then
    MsgBox ("��������� ������ ���� �� ������ �������")
    Exit Sub
End If

If wshWorkSheet.Cells(3, 2).Value < 0 Then
    MsgBox ("����� �� ����� ���� ������ 0%")
    Exit Sub
Else
    dTaxRate = wshWorkSheet.Cells(3, 2).Value / 100
End If

'theVariable(y, x)
'y is for Total, PrePay, PostPay
'x is for Cost, VAT, Cost+VAT
'Cost
dMoney(0, 0) = wshWorkSheet.Cells(2, 2).Value '123,456,789,012.3456789
dMoney(0, 0) = WorksheetFunction.Round(dMoney(0, 0), 2) '123,456,789,012.35
wshWorkSheet.Cells(7, 2).Value = dMoney(0, 0)
'VAT
dTaxRate = wshWorkSheet.Cells(3, 2).Value / 100
dMoney(0, 1) = dMoney(0, 0) * dTaxRate
dMoney(0, 1) = WorksheetFunction.Round(dMoney(0, 1), 2)
wshWorkSheet.Cells(7, 3).Value = dMoney(0, 1)
'Cost+VAT
dMoney(0, 2) = dMoney(0, 0) + dMoney(0, 1)
wshWorkSheet.Cells(7, 4).Value = dMoney(0, 2)

Call PaymentSchedule

End Sub

Private Sub SubVAT()
If wshWorkSheet.Cells(2, 2).Value > 999999999999.99 Then
    MsgBox ("��������� ������ ���� ������ ���������")
    Exit Sub
End If
If wshWorkSheet.Cells(2, 2).Value < 0.01 Then
    MsgBox ("��������� ������ ���� �� ������ �������")
    Exit Sub
End If

If wshWorkSheet.Cells(3, 2).Value < 0 Then
    MsgBox ("����� �� ����� ���� ������ 0%")
    Exit Sub
Else
    dTaxRate = wshWorkSheet.Cells(3, 2).Value / 100
End If

'theVariable(y, x)
'y is for Total, PrePay, PostPay
'x is for Cost, VAT, Cost+VAT
'Cost+VAT
dMoney(0, 2) = wshWorkSheet.Cells(2, 2).Value '123,456,789,012.3456789
dMoney(0, 2) = WorksheetFunction.Round(dMoney(0, 2), 2) '123,456,789,012.35
wshWorkSheet.Cells(7, 4).Value = dMoney(0, 2)
'Cost
dMoney(0, 0) = dMoney(0, 2) / (1 + dTaxRate)
dMoney(0, 0) = WorksheetFunction.Round(dMoney(0, 0), 2)
wshWorkSheet.Cells(7, 2).Value = dMoney(0, 0)
'VAT
dMoney(0, 1) = dMoney(0, 2) - dMoney(0, 0)
wshWorkSheet.Cells(7, 3).Value = dMoney(0, 1)

Call PaymentSchedule

End Sub

Private Sub PaymentSchedule()
'theVariable(y, x)
'y is for Total, PrePay, PostPay
'x is for Cost, VAT, Cost+VAT

'ForePay
If wshWorkSheet.Cells(4, 2).Value > 100 Then
    MsgBox ("����� �� ����� ���� ������ 100%")
    Exit Sub
ElseIf wshWorkSheet.Cells(4, 2).Value < 0 Then
    MsgBox ("����� �� ����� ���� ������ 0%")
    Exit Sub
Else
    dPayRate = wshWorkSheet.Cells(4, 2).Value / 100
End If
'Cost
dMoney(1, 0) = dMoney(0, 0) * dPayRate
dMoney(1, 0) = WorksheetFunction.Round(dMoney(1, 0), 2)
wshWorkSheet.Cells(8, 2).Value = dMoney(1, 0)
'VAT
dMoney(1, 1) = dMoney(0, 1) * dPayRate
dMoney(1, 1) = WorksheetFunction.Round(dMoney(1, 1), 2)
wshWorkSheet.Cells(8, 3).Value = dMoney(1, 1)
'Cost+VAT
dMoney(1, 2) = dMoney(0, 2) * dPayRate
dMoney(1, 2) = WorksheetFunction.Round(dMoney(1, 2), 2)
wshWorkSheet.Cells(8, 4).Value = dMoney(1, 2)

'PostPay
'Cost
dMoney(2, 0) = dMoney(0, 0) - dMoney(1, 0)
wshWorkSheet.Cells(9, 2).Value = dMoney(2, 0)
'VAT
dMoney(2, 1) = dMoney(0, 1) - dMoney(1, 1)
wshWorkSheet.Cells(9, 3).Value = dMoney(2, 1)
'Cost+VAT
dMoney(2, 2) = dMoney(0, 2) - dMoney(1, 2)
wshWorkSheet.Cells(9, 4).Value = dMoney(2, 2)

Call SpellTax

End Sub

Private Sub SpellTax()
Dim x As Byte, y As Byte

'theVariable(y, x)
'y is for Total, PrePay, PostPay
'x is for Cost, VAT, Cost+VAT
For y = 0 To 2
    For x = 0 To 2
        Call GetBillions(y, x)
        Call GetMillions(y, x)
        Call GetThousands(y, x)
        Call GetRoubles(y, x)
        Call GetKopecks(y, x)
    Next x

    'Integrate
    txtSpell = CVar( _
    LTrim(Format(dMoney(y, 0), _
    "### ### ### ##0.00")) & " (" & _
    SpellAll(y, 0) & _
    ", ����� ����, ��� �� ������ " & _
    wshWorkSheet.Cells(3, 2).Value & _
    "% � ������� " & _
    LTrim(Format(dMoney(y, 1), _
    "### ### ### ##0.00")) & " (" & _
    SpellAll(y, 1) & _
    ", ����� � ��� " & _
    LTrim(Format(dMoney(y, 2), _
    "### ### ### ##0.00")) & " (" & _
    SpellAll(y, 2) & _
    ".")

    'Set "Alt+160" Non-breaking Spaces After Digits
    For i = 0 To 9
        txtSpell = Replace(txtSpell, i & " ", i & Chr(160))
    Next i
    txtSpell = Replace(txtSpell, Chr(160) & "(", " (")
    
    'Remove Unnecessary Spaces
    txtSpell = Replace(txtSpell, "( ", "(")
    txtSpell = Replace(txtSpell, " )", ")")
    txtSpell = Replace(txtSpell, "  ", " ")
    
    'Capitalize Letters
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    txtSpell = Replace(txtSpell, "(�", "(�")
    
    'Show the Result
    wshWorkSheet.Cells(y * 7 + 12, 1).Value = txtSpell
Next y

CopyText (12) 'y = 12

End Sub

Function GetBillions(y As Byte, x As Byte)
    Dim n As Double
    n = dMoney(y, x) '123,456,789,012.35
    n = n / 1000000000 '123.45678901235
    n = Fix(n) '123 'Remove Last Digits
    n = n / 1000 '0.123
    iBillions(y, x) = n * 1000 '123
End Function

Function GetMillions(y As Byte, x As Byte)
    Dim n As Double
    n = dMoney(y, x) '123,456,789,012.35
    n = n / 1000000 '123,456.78901235
    n = Fix(n) '123,456 'Remove Last Digits
    n = n / 1000 '123.456
    n = n - Fix(n) '0.456
    iMillions(y, x) = n * 1000 '456
End Function

Function GetThousands(y As Byte, x As Byte)
    Dim n As Double
    n = dMoney(y, x) '123,456,789,012.35
    n = n / 1000 '123,456,789.01235
    n = Fix(n) '123,456,789 'Remove Last Digits
    n = n / 1000 '123,456.789
    n = n - Fix(n) '0.789
    iThousands(y, x) = n * 1000 '789
End Function

Function GetRoubles(y As Byte, x As Byte)
    Dim n As Double
    n = dMoney(y, x) '123,456,789,012.35
    n = Fix(n) '123,456,789,012 'Remove Last Digits
    n = n / 1000 '123,456,789.012
    n = n - Fix(n) '0.012
    iRoubles(y, x) = n * 1000 '12
End Function

Function GetKopecks(y As Byte, x As Byte)
    Dim n As Double
    n = dMoney(y, x) '123,456,789,012.3456
    n = Fix(n) '123,456,789,012
    n = dMoney(y, x) - n '0.3456
    n = n * 100 '34.56
    iKopecks(y, x) = WorksheetFunction.Round(n, 2) '35
End Function

Function SpellAll(y As Byte, x As Byte) As String
If Fix(dMoney(y, x)) = 0 Then
    SpellAll = "���� ������ "
Else
    If iBillions(y, x) > 0 Then
        SpellAll = SpellBillions(iBillions(y, x))
    End If
    If iMillions(y, x) > 0 Then
        SpellAll = SpellAll + SpellMillions(iMillions(y, x))
    End If
    If iThousands(y, x) > 0 Then
        SpellAll = SpellAll + SpellThousands(iThousands(y, x))
    End If
    SpellAll = SpellAll + SpellRoubles(iRoubles(y, x))
End If
    SpellAll = SpellAll + SpellKopecks(iKopecks(y, x))
End Function

Function SpellBillions(iFragment As Integer) As String
Dim n As Double
Dim iTens As Integer
Dim bLastDigit As Byte

n = iFragment / 100 ' 1.23
n = n - Fix(n) ' 1.23 - 1.00 = 0.23
iTens = CInt(n * 100) ' 23
n = iFragment / 10 '12.3
n = n - Fix(n) '12.3 - 12.0 = 0.3
bLastDigit = CByte(n * 10) '3

SpellBillions = _
SpellHundreds(iFragment) & _
SpellTens(iTens, bLastDigit, "Nominativus", "Masculinum")

If iTens > 4 And iTens < 21 Then
    SpellBillions = SpellBillions + " ����������"
ElseIf bLastDigit = 0 Then
    SpellBillions = SpellBillions + " ����������"
ElseIf bLastDigit = 1 Then
    SpellBillions = SpellBillions + " ��������"
ElseIf bLastDigit < 5 Then
    SpellBillions = SpellBillions + " ���������"
Else
    SpellBillions = SpellBillions + " ����������"
End If
End Function

Function SpellMillions(iFragment As Integer) As String
Dim n As Double
Dim iTens As Integer
Dim bLastDigit As Byte

n = iFragment / 100 ' 1.23
n = n - Fix(n) ' 1.23 - 1.00 = 0.23
iTens = CInt(n * 100) ' 23
n = iFragment / 10 '12.3
n = n - Fix(n) '12.3 - 12.0 = 0.3
bLastDigit = CByte(n * 10) '3

SpellMillions = _
SpellHundreds(iFragment) & _
SpellTens(iTens, bLastDigit, "Nominativus", "Masculinum")

If iTens > 4 And iTens < 21 Then
    SpellMillions = SpellMillions + " ���������"
ElseIf bLastDigit = 0 Then
    SpellMillions = SpellMillions + " ���������"
ElseIf bLastDigit = 1 Then
    SpellMillions = SpellMillions + " �������"
ElseIf bLastDigit < 5 Then
    SpellMillions = SpellMillions + " ��������"
Else
    SpellMillions = SpellMillions + " ���������"
End If
End Function

Function SpellThousands(iFragment As Integer) As String
Dim n As Double
Dim iTens As Integer
Dim bLastDigit As Byte

n = iFragment / 100 ' 1.23
n = n - Fix(n) ' 1.23 - 1.00 = 0.23
iTens = CInt(n * 100) ' 23
n = iFragment / 10 '12.3
n = n - Fix(n) '12.3 - 12.0 = 0.3
bLastDigit = CByte(n * 10) '3

SpellThousands = _
SpellHundreds(iFragment) & _
SpellTens(iTens, bLastDigit, "Nominativus", "Femininum")

If iTens > 4 And iTens < 21 Then
    SpellThousands = SpellThousands + " �����"
ElseIf bLastDigit = 0 Then
    SpellThousands = SpellThousands + " �����"
ElseIf bLastDigit = 1 Then
    SpellThousands = SpellThousands + " ������"
ElseIf bLastDigit < 5 Then
    SpellThousands = SpellThousands + " ������"
Else
    SpellThousands = SpellThousands + " �����"
End If
End Function

Function SpellRoubles(iFragment As Integer) As String
Dim n As Double
Dim iTens As Integer
Dim bLastDigit As Byte

n = iFragment / 100 ' 1.23
n = n - Fix(n) ' 1.23 - 1.00 = 0.23
iTens = CInt(n * 100) ' 23
n = iFragment / 10 '12.3
n = n - Fix(n) '12.3 - 12.0 = 0.3
bLastDigit = CByte(n * 10) '3

SpellRoubles = _
SpellHundreds(iFragment) & _
SpellTens(iTens, bLastDigit, "Nominativus", "Masculinum")

If iTens > 4 And iTens < 21 Then
    SpellRoubles = SpellRoubles + " ������ "
ElseIf bLastDigit = 0 Then
    SpellRoubles = SpellRoubles + " ������ "
ElseIf bLastDigit = 1 Then
    SpellRoubles = SpellRoubles + " ����� "
ElseIf bLastDigit < 5 Then
    SpellRoubles = SpellRoubles + " ����� "
Else
    SpellRoubles = SpellRoubles + " ������ "
End If
End Function

Function SpellHundreds(iFragment As Integer) As String
If iFragment < 100 Then
    SpellHundreds = ""
ElseIf iFragment < 200 Then
    SpellHundreds = " ���"
ElseIf iFragment < 300 Then
    SpellHundreds = " ������"
ElseIf iFragment < 400 Then
    SpellHundreds = " ������"
ElseIf iFragment < 500 Then
    SpellHundreds = " ���������"
ElseIf iFragment < 600 Then
    SpellHundreds = " �������"
ElseIf iFragment < 700 Then
    SpellHundreds = " ��������"
ElseIf iFragment < 800 Then
    SpellHundreds = " �������"
ElseIf iFragment < 900 Then
    SpellHundreds = " ���������"
ElseIf iFragment < 1000 Then
    SpellHundreds = " ���������"
Else
    Call ClearAll
    MsgBox ("������ � ������")
End If
End Function

Function SpellTens(iTens As Integer, bLastDigit As Byte, CasusGrammaticus As String, Genus As String) As String
If iTens < 10 Then
    If Genus = "Femininum" Then
        SpellTens = SpellOnesFemininum(bLastDigit)
    Else
        SpellTens = SpellOnesMasculinum(bLastDigit)
    End If
    Exit Function
ElseIf iTens = 10 Then
    SpellTens = " ������"
    Exit Function
ElseIf iTens = 11 Then
    SpellTens = " �����������"
    Exit Function
ElseIf iTens = 12 Then
    SpellTens = " ����������"
    Exit Function
ElseIf iTens = 13 Then
    SpellTens = " ����������"
    Exit Function
ElseIf iTens = 14 Then
    SpellTens = " ������������"
    Exit Function
ElseIf iTens = 15 Then
    SpellTens = " ����������"
    Exit Function
ElseIf iTens = 16 Then
    SpellTens = " �����������"
    Exit Function
ElseIf iTens = 17 Then
    SpellTens = " ����������"
    Exit Function
ElseIf iTens = 18 Then
    SpellTens = " ������������"
    Exit Function
ElseIf iTens = 19 Then
    SpellTens = " ������������"
    Exit Function
ElseIf iTens < 30 Then
    SpellTens = " ��������"
ElseIf iTens < 40 Then
    SpellTens = " ��������"
ElseIf iTens < 50 Then
    SpellTens = " �����"
ElseIf iTens < 60 Then
    SpellTens = " ���������"
ElseIf iTens < 70 Then
    SpellTens = " ����������"
ElseIf iTens < 80 Then
    SpellTens = " ���������"
ElseIf iTens < 90 Then
    SpellTens = " �����������"
Else
    SpellTens = " ���������"
End If

If Genus = "Femininum" Then
    SpellTens = SpellTens + SpellOnesFemininum(bLastDigit)
Else
    SpellTens = SpellTens + SpellOnesMasculinum(bLastDigit)
End If
End Function

Function SpellOnesFemininum(bLastDigit As Byte) As String
If bLastDigit = 1 Then
    SpellOnesFemininum = " ����"
ElseIf bLastDigit = 2 Then
    SpellOnesFemininum = " ���"
Else
    SpellOnesFemininum = SpellOnesNeutrum(bLastDigit)
End If
End Function

Function SpellOnesMasculinum(bLastDigit As Byte) As String
If bLastDigit = 1 Then
    SpellOnesMasculinum = " ����"
ElseIf bLastDigit = 2 Then
    SpellOnesMasculinum = " ���"
Else
    SpellOnesMasculinum = SpellOnesNeutrum(bLastDigit)
End If
End Function

Function SpellOnesNeutrum(bLastDigit As Byte) As String
If bLastDigit = 3 Then
    SpellOnesNeutrum = " ���"
ElseIf bLastDigit = 4 Then
    SpellOnesNeutrum = " ������"
ElseIf bLastDigit = 5 Then
    SpellOnesNeutrum = " ����"
ElseIf bLastDigit = 6 Then
    SpellOnesNeutrum = " �����"
ElseIf bLastDigit = 7 Then
    SpellOnesNeutrum = " ����"
ElseIf bLastDigit = 8 Then
    SpellOnesNeutrum = " ������"
ElseIf bLastDigit = 9 Then
    SpellOnesNeutrum = " ������"
End If
End Function

Function SpellKopecks(iTens As Integer) As String
Dim n As Double
Dim bLastDigit As Byte

n = iTens / 10 '1.2
n = n - Fix(n) '1.2 - 1.0 = 0.2
bLastDigit = CByte(n * 10) '2

'Spell
If iTens = 0 Then
    SpellKopecks = "00" & Chr(160) & "������)"
ElseIf iTens = 1 Then
    SpellKopecks = "01" & Chr(160) & "�������)"
ElseIf iTens < 5 Then
    SpellKopecks = "0" & CStr(iTens) & Chr(160) & "�������)"
ElseIf iTens < 10 Then
    SpellKopecks = "0" & CStr(iTens) & Chr(160) & "������)"
ElseIf iTens < 21 Then
    SpellKopecks = CStr(iTens) & Chr(160) & "������)"
ElseIf bLastDigit = 0 Then
    SpellKopecks = CStr(iTens) & Chr(160) & "������)"
ElseIf bLastDigit = 1 Then
    SpellKopecks = CStr(iTens) & Chr(160) & "�������)"
ElseIf bLastDigit < 5 Then
    SpellKopecks = CStr(iTens) & Chr(160) & "�������)"
Else
    SpellKopecks = CStr(iTens) & Chr(160) & "������)"
End If
End Function

Private Sub ClearAll()
    wshWorkSheet.Range("B2").Value = ""
    wshWorkSheet.Range("B7:D9").Value = ""
    wshWorkSheet.Range("A12:F16").Value = ""
    wshWorkSheet.Range("A19:F23").Value = ""
    wshWorkSheet.Range("A26:F30").Value = ""
End Sub

Private Sub CopyTotal()
    CopyText (12)
End Sub

Private Sub CopyPrePay()
    CopyText (19)
End Sub

Private Sub CopyPostPay()
    CopyText (26)
End Sub

Private Sub CopyText(y As Byte)
    With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'Magic Number
        .SetText wshWorkSheet.Cells(y, 1).Value 'txtSpell
        .PutInClipboard
    End With
    MsgBox ("����� ���������� � ����� ������." & Chr(13) & "������ ��������� � Word (Ctrl+V)")
End Sub

