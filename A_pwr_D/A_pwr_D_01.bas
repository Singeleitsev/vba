Attribute VB_Name = "??1"
Sub A_pwr_D()

A = 2 'Argument
D = 100 'Degree

L = 1 'Length
With ActiveSheet
    .Cells(1, 1) = A
    For y = 2 To D
        For x = L To 1 Step -1
            n = .Cells(y, x) + .Cells(y - 1, x) * A
            If n < 1000 Then
                .Cells(y, x) = n
            Else
                r = Int(n / 1000)
                .Cells(y, x) = n - r * 1000
                If x > 1 Then
                    .Cells(y, x - 1) = r
                Else
                    L = L + 1
                    Columns(1).Insert
                    .Cells(y, x) = r
                End If
            End If
        Next x
    Next y
    For x = 1 To L
        .Columns(x).AutoFit
    Next x
    For y = 1 To D
        For x = 2 To L
            If .Cells(y, x - 1) <> "" Then
                .Cells(y, x).NumberFormat = "000"
            End If
        Next x
    Next y
End With
End Sub
