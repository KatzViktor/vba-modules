Attribute VB_Name = "Effects"
' ��������� � ���������
If VesZadachi >= 25 Or Sheets("����1").Cells(i, 5).Value = 2 Or Sheets("����1").Cells(i, 6).Value = 4 Then
    Sheets("����1").Cells(i, 3).Value = "�"
    Sheets("����1").Cells(i, 3).HorizontalAlignment = xlCenter
    Sheets("����1").Cells(i, 3).VerticalAlignment = xlCenter
    Sheets("����1").Range(Cells(i, 1), Cells(i, 8)).Interior.Color = RGB(240, 250, 114)
Else
    Sheets("����1").Cells(i, 3).Value = ""
    Sheets("����1").Cells(i, 3).HorizontalAlignment = xlCenter
    Sheets("����1").Cells(i, 3).VerticalAlignment = xlCenter
    Sheets("����1").Range(Cells(i, 1), Cells(i, 8)).Interior.Color = RGB(255, 255, 255)
End If
''''''
