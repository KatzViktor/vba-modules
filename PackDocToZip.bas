
Sub ZipVbs(zipfile, inputfile)

CreateObject("Scripting.FileSystemObject").CreateTextFile(zipfile, True).Write "PK" & Chr$(5) & Chr$(6) & String(18, vbNullChar)
Dim objShell
Set objShell = CreateObject("Shell.Application")
objShell.Namespace(zipfile).CopyHere (inputfile)

End Sub

Sub Zippy()

Dim NewPath, spath, zippath, zfile As String
Dim fso As Object

spath = Sheets("Main").Cells(8, 5).Value & _
          Sheets("Main").Cells(3, 9).Value & "\" & _
          "Новый документ" & Sheets("Main").Cells(3, 7).Value & ".doc"
Set objShell = CreateObject("Shell.Application")
objShell.Explore (spath)
zippath = Sheets("Main").Cells(8, 5).Value & _
          Sheets("Main").Cells(3, 9).Value & "\" & _
          "Новый документ" & Sheets("Main").Cells(3, 7).Value & ".zip"
zfile = spath
ZipVbs zippath, zfile

NewPath = Sheets("Main").Cells(11, 5).Value & _
          "Документ номер " & Sheets("Main").Cells(3, 7).Value & ".doc"

Set fso = CreateObject("Scripting.FileSystemObject")
If Dir(spath) <> "" Then ' если файл существует копируем
    fso.CopyFile spath, NewPath
    fso.DeleteFile spath
End If

End Sub

