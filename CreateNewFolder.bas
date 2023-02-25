Attribute VB_Name = "NewFolder"
Sub CreateNewFolder()
Dim fso As Object, i As Integer
Dim FolderPath As String

FolderPath = Sheets("Main").Cells(15, 5).Value & _
          Sheets("Main").Cells(3, 9).Value

Set fso = CreateObject("Scripting.FileSystemObject")
With fso
    .CreateFolder (FolderPath)
End With

End Sub


