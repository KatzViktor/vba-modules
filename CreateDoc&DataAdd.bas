
Sub CreateNewFile()
Dim fso As Object, i As Integer
Dim OldPath, NewPath, DiskPath As String
Dim DocFullName, WA As Object
Dim oMyDoc As Object
Dim SZagolovok, SDate, SIzh, SOD As String

Set fso = CreateObject("Scripting.FileSystemObject")
' копирование файла

OldPath = Sheets("Main").Cells(16, 5).Value & _
          "1700 " & Sheets("Main").Cells(3, 6).Value & ".doc"
DiskPath = Sheets("Main").Cells(8, 5).Value & _
          Sheets("Main").Cells(3, 9).Value & "\" & _
          "1700 " & Sheets("Main").Cells(3, 7).Value & ".doc"
If Dir(OldPath) <> "" Then ' если файл существует копируем
    fso.CopyFile OldPath, DiskPath
End If

' конец блока копирования

SZagolovok = "(по состоянию на 17.00 " & Sheets("Main").Cells(5, 7).Value & ")"
SDate = Sheets("Main").Cells(4, 7).Value
SIzh = Sheets("Main").Cells(3, 3).Value & Sheets("Main").Cells(16, 2).Value
SOD = "Автор " & Sheets("Main").Cells(4, 4).Value & " " & Sheets("Main").Cells(4, 5).Value

DocFullName = DiskPath
Set WA = CreateObject("Word.Application")
WA.Visible = False
Set oMyDoc = WA.Documents.Open(DocFullName)

oMyDoc.Range(oMyDoc.Paragraphs(11).Range.Start, oMyDoc.Paragraphs(11).Range.End - 1).Text = SZagolovok
oMyDoc.Tables(1).Cell(1, 1).Range.Text = SIzh
oMyDoc.Tables(1).Cell(1, 2).Range.Text = SDate
oMyDoc.Tables(1).Cell(1, 3).Range.Text = SOD

oMyDoc.Close
WA.Quit False
Set oMyDoc = Nothing: Set WA = Nothing

End Sub
