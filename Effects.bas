Attribute VB_Name = "Effects"
' выровнять и закрасить
If VesZadachi >= 25 Or Sheets("Лист1").Cells(i, 5).Value = 2 Or Sheets("Лист1").Cells(i, 6).Value = 4 Then
    Sheets("Лист1").Cells(i, 3).Value = "С"
    Sheets("Лист1").Cells(i, 3).HorizontalAlignment = xlCenter
    Sheets("Лист1").Cells(i, 3).VerticalAlignment = xlCenter
    Sheets("Лист1").Range(Cells(i, 1), Cells(i, 8)).Interior.Color = RGB(240, 250, 114)
Else
    Sheets("Лист1").Cells(i, 3).Value = ""
    Sheets("Лист1").Cells(i, 3).HorizontalAlignment = xlCenter
    Sheets("Лист1").Cells(i, 3).VerticalAlignment = xlCenter
    Sheets("Лист1").Range(Cells(i, 1), Cells(i, 8)).Interior.Color = RGB(255, 255, 255)
End If
''''''
