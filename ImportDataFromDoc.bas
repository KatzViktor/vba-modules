
Sub ImportDataFromDoc()

Dim MyPos As Integer
Dim Element As String
Dim objFSO, sourcefolder, objFile
Dim spath, sname
Dim r
Dim S, DocFullName, WA As Object
Dim oMyDoc As Object
Dim NomerStroki, NomerStolbca As Integer

NomerStroki = 2
NomerStolbca = 1
spath = "D:\Data.doc"
    Do While Sheets("Данные").Cells(NomerStroki, 2) <> ""
            'NomerStroki = NomerStroki + 1
            If Sheets("Данные").Cells(NomerStroki, 3) = "" Then
                Element = Sheets("Данные").Cells(NomerStroki, 2)
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                Set sourcefolder = objFSO.getfolder(spath)
                'r = 1
                    For Each fileitem In sourcefolder.Files
                        'r = r + 1
                        objBaseName = objFSO.GetBaseName(fileitem)
                        Set objFile = objFSO.GetFile(fileitem)
                        MyPos = InStr(1, objBaseName, Element, 1)
                        If (MyPos = 5) Then Sheets("Данные").Cells(NomerStroki, 5).Value = objFile.ParentFolder & "\" & objBaseName & ".doc"
                    Next fileitem
            
            DocFullName = Sheets("Данные").Cells(NomerStroki, 5)
            Set WA = CreateObject("Word.Application")
            WA.Visible = False
            Set oMyDoc = WA.Documents.Open(DocFullName)
                For NomerStolbca = 2 To 4
                    S = oMyDoc.Tables(1).Cell(2, NomerStolbca).Range.Text
                    S = Replace(S, Chr(7), "") 'удаление символа конца ячейки
                    S = Left(S, Len(S) - 1)
                    Sheets("Данные").Cells(NomerStroki, NomerStolbca) = S
                Next NomerStolbca
                    S = oMyDoc.Tables(1).Cell(3, 1).Range.Text
                    S = Replace(S, Chr(7), "") 'удаление символа конца ячейки
                    S = Left(S, Len(S) - 1)
                    Sheets("Данные").Cells(NomerStroki, 5) = S
                For NomerStolbca = 6 To 9
                    S = oMyDoc.Tables(1).Cell(2, NomerStolbca).Range.Text
                    S = Replace(S, Chr(7), "") 'удаление символа конца ячейки
                    S = Left(S, Len(S) - 1)
                    Sheets("Данные").Cells(NomerStroki, NomerStolbca) = S
                Next NomerStolbca
                
            oMyDoc.Close 0
            WA.Quit False
            Set oMyDoc = Nothing: Set WA = Nothing
        End If
    NomerStroki = NomerStroki + 1
    Loop
End Sub

