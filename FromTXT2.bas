Attribute VB_Name = "Module6"
Function Find(ByVal iC1 As Integer, sFS1 As String, aMass() As String) As Integer
    Do Until InStr(aMass(iC1), sFS1) <> 0
        iC1 = iC1 + 1
    Loop
    Find = iC1
End Function
Function DelSpace(ByVal sPhrase As String) As String
    sPhrase = Trim(sPhrase) 'удаление пробелов слева и справа
        Do While InStr(sPhrase, "  ")
            sPhrase = Replace(sPhrase, "  ", " ") 'замена двойных пробелов одинарными
        Loop
    sPhrase = Right(sPhrase, Len(sPhrase) - 1) 'удаление одного символа слева
    sPhrase = Left(sPhrase, Len(sPhrase) - 1) 'удаление одного симовола справа
End Function


'Sub Test()

Dim sPathOfTXT As String, sData As String
Dim iQuantOfLines As Integer, iStartCount As Integer, iEndCount As Integer, iNextString As Integer, iCounter As Integer
Dim aTXT() As String, aMass1(40, 2) As String, sFindSymb As String
Dim iFindDate As Integer, iUmenshStroki As Integer, sDFT As String, Poopi As String
Dim aArrayOfData() As String, iQntOfPlane As Integer
Dim iIII As Integer, iJJJJ As Integer, XoX As Integer
Dim sPut As String, sFormat As String, sNazvaniee As String
Dim iPosOfYear As Integer

'______________________________________________________________________________________________________
'блок копирования документа в массив
'______________________________________________________________________________________________________
  sPut = "D:\Общее\"
  sFormat = ".txt"
  XoX = 1
  Do While Len(Sheets("1").Cells(XoX, 1).Value) > 0
  iQuantOfLines = 0
  iStartCount = 0
  iEndCount = 0
  iNextString = 0
  iCounter = 0
  iFindDate = 0
  iUmenshStroki = 0
  iQntOfPlane = 0
  iIII = 0
  iJJJJ = 0
  iPosOfYear = 0
  
    sNazvaniee = Sheets("1").Cells(XoX, 1).Value
  
    sPathOfTXT = sPut & sNazvaniee & sFormat
        Open sPathOfTXT For Input As #1
        iQuantOfLines = 0   'последующий цикл осуществляет подсчет строк в исходном документе
            Do Until EOF(1)
                Line Input #1, sData
                iQuantOfLines = iQuantOfLines + 1
            Loop
        Close #1
    ReDim aTXT(iQuantOfLines) 'задается размерность массива для копирования в него документа
        Open sPathOfTXT For Input As #1
        iQuantOfLines = 0   'последующий цикл осуществляет подсчет строк в исходном документе
            Do Until EOF(1)
                Line Input #1, aTXT(iQuantOfLines)
                iQuantOfLines = iQuantOfLines + 1
            Loop
        Close #1
'___________________________________________________________________________________________________
'блок добавления задачи
'________________________________________________________________________________________________

   
        sFindSymb = "I-"
        Do Until InStr(aTXT(iStartCount), sFindSymb) <> 0 ' КОМПАНОВКА задачи
            sData = aTXT(iStartCount)
                Do While InStr(sData, "  ")
                    sData = Replace(sData, "  ", " ") 'замена двойных пробелов одинарными
                Loop
            sData = Right(sData, Len(sData) - 1) 'удаление одного символа слева
            sData = Left(sData, Len(sData) - 1) 'удаление одного симовола справа
            sData = Trim(sData) 'удаление пробелов слева и справа
            aMass1(iMass1X, 0) = aMass1(iMass1X, 0) + " " + sData
            iStartCount = iStartCount + 1
        Loop
'конец блока добавления цели
iStartCount = iStartCount + 1

'начало блока добавления типов
        Poopi = aMass1(iMass1X, 0)
        sFindSymb = "I="
        Do Until InStr(aTXT(iStartCount), sFindSymb) <> 0
            sData = aTXT(iStartCount)
            iDemention = InStr(sData, ":") + 1 'сокращение строки
            sData = Right(sData, iDemention) 'сокращение строки
           
            sDFT = Trim(Left(sData, InStr(sData, ":")))  'выборка данных из таблицы
            If sDFT <> "" Then
                
                aMass1(iMass1X, 1) = Trim(Left(sData, InStr(sData, ":") - 1)) 'добавление типа
                iDemention = Len(sData) - 8 'сокращение строки
                sData = Right(sData, iDemention) 'сокращение строки
                aMass1(iMass1X, 2) = Trim(Left(sData, 4))
                aMass1(iMass1X, 0) = Poopi
                
               
                iMass1X = iMass1X + 1
            End If
        'iMass1X = iMass1X + 1
        iStartCount = iStartCount + 1
        Loop
'___________________________________________
Do While iStartCount + 1 < iEndCount
   iStartCount = iStartCount + 1
    sData = aTXT(iStartCount)
        sFindSymb = "I-"
        Do Until InStr(aTXT(iStartCount), sFindSymb) <> 0
            sData = aTXT(iStartCount)
                Do While InStr(sData, "  ")
                    sData = Replace(sData, "  ", " ") 'замена двойных пробелов одинарными
                Loop
            sData = Right(sData, Len(sData) - 1) 'удаление одного символа слева
            sData = Left(sData, Len(sData) - 1) 'удаление одного симовола справа
            sData = Trim(sData) 'удаление пробелов слева и справа
            aMass1(iMass1X, 0) = aMass1(iMass1X, 0) + " " + sData
            iStartCount = iStartCount + 1
        Loop
'конец блока добавления задачи
iStartCount = iStartCount + 1

'начало блока добавления типов
        Poopi = aMass1(iMass1X, 0)
        sFindSymb = "I="
        Do Until InStr(aTXT(iStartCount), sFindSymb) <> 0
            sData = aTXT(iStartCount)
        sKonec = Left(sData, 2)
        If sKonec <> "==" Then
            iDemention = InStr(sData, ":") + 1 'сокращение строки
            sData = Right(sData, iDemention) 'сокращение строки
            sDFT = Trim(Left(sData, InStr(sData, ":")))
            If sDFT <> "" Then

                aMass1(iMass1X, 1) = Trim(Left(sData, InStr(sData, ":") - 1)) 'добавление типа
                iDemention = Len(sData) - 8 'сокращение строки
                sData = Right(sData, iDemention) 'сокращение строки
                aMass1(iMass1X, 2) = Trim(Left(sData, 4))
                aMass1(iMass1X, 0) = Poopi
               
                iMass1X = iMass1X + 1
            End If
            Else:
            iStartCount = iStartCount + 2
             iMass1X = iMass1X - 1
        End If
        iMass1X = iMass1X + 1
        iStartCount = iStartCount + 1
        
        Loop


'__________________________________________


Loop
iJach = 1

Do While Sheets("Лист1").Cells(iJach, 11).Value <> ""
    iJach = iJach + 1
Loop


iQntOfRow = 0

iEndCount = iMass1X + 1
Do While iMass1X > 0
  
 
  sData = aMass1(iQntOfRow, 2)
  If Len(sData) > 0 Then
     Else:
        
        iStartCount = iQntOfRow
        iNextString = iStartCount + 1
            Do While iNextString < iEndCount
                aMass1(iStartCount, 0) = aMass1(iNextString, 0)
                aMass1(iStartCount, 1) = aMass1(iNextString, 1)
                aMass1(iStartCount, 2) = aMass1(iNextString, 2)
                iNextString = iNextString + 1
                iStartCount = iStartCount + 1
            Loop
 
        iEndCount = iEndCount - 1
        iQntOfRow = iQntOfRow - 1
  End If
iQntOfRow = iQntOfRow + 1
iMass1X = iMass1X - 1
Debug.Print aMass1(iQntOfRow, 0) & aMass1(iQntOfRow, 1) & aMass1(iQntOfRow, 2)
Loop
iMass1X = iEndCount
iQntOfRow = 0


Do While iMass1X > 0

    If InStr(aMass1(iQntOfRow, 1), "MQ") > 0 Then
    
    Else
    If InStr(aMass1(iQntOfRow, 1), "МQ") > 0 Then
    Else
    iQntOfPlane = aMass1(iQntOfRow, 2)
        If iQntOfPlane > 1 Then
            Do While iQntOfPlane > 0
                Sheets("Полеты").Cells(iJach, 12).Value = aMass1(iQntOfRow, 0)
                iJach = iJach + 1
                iQntOfPlane = iQntOfPlane - 1
            Loop
            Else
                Sheets("Полеты").Cells(iJach, 12).Value = aMass1(iQntOfRow, 0)
                iJach = iJach + 1
        End If
     End If
     End If



iMass1X = iMass1X - 1
iQntOfRow = iQntOfRow + 1
Loop


'______________________________________________________________________________________
'НАЧАЛО АНАЛИЗА Задач
'_______________________________________________________________________________________
iIII = 2
Do While Sheets("Полеты").Cells(iIII, 12).Value <> ""

If InStr(Sheets("Полеты").Cells(iIII, 12).Value, "КОНТРОЛЬ") Then
Sheets("Лист1").Cells(iIII, 11).Value = "Контроль"
End If

If InStr(Sheets("Полеты").Cells(iIII, 12).Value, "ПРОПУСК") Then
Sheets("Лист1").Cells(iIII, 11).Value = "Пропуск"
End If

iIII = iIII + 1
Loop
'_____________________________________________________________________________________________
'КОНЕЦ АНАЛИЗА Задач
'_____________________________________________________________________________________________
XoX = XoX + 1
Loop
End Sub
