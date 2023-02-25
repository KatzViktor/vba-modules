Attribute VB_Name = "Module3"
Function TrueSeparate(sStroka As String)
Dim sMassiv(40) As String, sYear As String
Dim iI As Integer, iJ As Integer, iKaretka As Integer, iCounter As Integer
Dim iDemention As Integer, sDate As String

sDate = Left(sStroka, 10)
iDemention = Len(sStroka) - 10 'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки


sMassiv(0) = Left(sStroka, InStr(sStroka, " ") - 1) 'Добавление типа
iDemention = Len(sStroka) - InStr(sStroka, " ") - 1 'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(1) = Left(sStroka, InStr(sStroka, " ") - 2) 
    If sMassiv(1) <> "Н/У" Then  '
        If InStr(sMassiv(1), "-") > 0 Then
            iSep = Len(sMassiv(1)) - InStr(sMassiv(1), "-")
            sMassiv(2) = Right(sMassiv(1), iSep)
            sMassiv(1) = Left(sMassiv(1), InStr(sMassiv(1), "-") - 1)
        Else 'если не содержит дефиса, то добавляется в индекс
            sMassiv(2) = sMassiv(1)
            sMassiv(1) = ""
        End If
    End If                          'конец добавления индекса
iDemention = Len(sStroka) - InStr(sStroka, " ") 'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(3) = Left(sStroka, InStr(sStroka, "(") - 2) 'добавление аэропорта вылета
iDemention = Len(sStroka) - InStr(sStroka, "(")  'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(4) = Left(sStroka, InStr(sStroka, ")") - 1) 'добавление даты и времени
    If InStr(sMassiv(4), " ") > 0 Then 'анализ даты времени
        iSep = Len(sMassiv(4)) - InStr(sMassiv(4), " ")
        sMassiv(5) = Right(sMassiv(4), iSep)
        sMassiv(4) = Left(sMassiv(4), InStr(sMassiv(4), " ") - 1)
        sYear = Right(sDate, 5)
        sMassiv(5) = sMassiv(5) + sYear
    Else
        sMassiv(5) = sDate 'если дата не стоит по умолчанию ставится дата за предыдущий день
    End If
iDemention = Len(sStroka) - InStr(sStroka, "-") - 1 'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(6) = Left(sStroka, InStr(sStroka, " - ") - 1) 'добавление района
iDemention = Len(sStroka) - InStr(sStroka, " - ") - 2 'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(7) = Left(sStroka, InStr(sStroka, "(") - 2) 'добавление аэропорта вылета
iDemention = Len(sStroka) - InStr(sStroka, "(")  'сокращение строки
sStroka = Right(sStroka, iDemention) 'сокращение строки
sMassiv(8) = Left(sStroka, InStr(sStroka, ")") - 1) 'добавление даты и времени
    If InStr(sMassiv(8), " ") > 0 Then 'анализ даты времени
        iSep = Len(sMassiv(8)) - InStr(sMassiv(8), " ")
        sMassiv(9) = Right(sMassiv(8), iSep)
        sMassiv(8) = Left(sMassiv(8), InStr(sMassiv(8), " ") - 1)
        sYear = Right(sDate, 5)
        sMassiv(9) = sMassiv(9) + sYear
    Else
        sMassiv(9) = sDate 'если дата не стоит по умолчанию ставится текущая дата
    End If

iI = 1
iJ = 1
Do While Sheets("СРА").Cells(iI, iJ).Value <> ""
    iI = iI + 1
Loop
Do While iJ <> 11
iCounter = iJ - 1
Sheets("СРА").Cells(iI, iJ).Value = sMassiv(iCounter)
iJ = iJ + 1
Loop
Sheets("Вылет").Cells(iI, 13).Value = "По данным Scramble"
End Function
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


Sub Test()

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
  sPut = "D:\Общее\Статистика\1\"
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
'______________________________________________________________________________________________________
'конец блока копирования документа
'______________________________________________________________________________________________________


'______________________________________________________________________________________________________
'начало блока получения даты
'______________________________________________________________________________________________________
    
    iFindDate = 1
    iFindDate = Find(iFindDate, "Начало", aTXT)
    iFindDate = iFindDate + 4
    aTXT(iFindDate) = Trim(aTXT(iFindDate))
    Poopi = aTXT(iFindDate)
    Poopi = Right(Poopi, 11)
    Poopi = Left(Poopi, 6)
    sDFT = Poopi & 2020
'______________________________________________________________________________________________________
'конец блока получения даты
'______________________________________________________________________________________________________


'____________________________________________________________________________________________________
'начало блока обработки характеристик
'____________________________________________________________________________________________________
    iStartCount = 0 'поиск начала раздела с характеристикой
    iStartCount = Find(iStartCount, "Начало", aTXT)
    iStartCount = iStartCount + 1
    iEndCount = 0 'поиск конца раздела с характеристикой
    iEndCount = Find(iEndCount, "Конец", aTXT)
    iEndCount = iEndCount - 2
        Do While iStartCount < iEndCount
            iNextString = iStartCount + 1
            
                If Left(aTXT(iNextString), InStr(aTXT(iNextString), " ") - 1) <> "-" Then 'если следующая строка является
                    aTXT(iStartCount) = aTXT(iStartCount) + aTXT(iNextString)             'частью предыдущей, то обе строки
                        Do While iNextString <> iQuantOfLines                             'объедняются, а массив смещается
                            iCounter = iNextString                                        'на одну строку
                            iNextString = iNextString + 1
                            aTXT(iCounter) = aTXT(iNextString)
                        Loop
                    iQuantOfLines = iQuantOfLines - 1
                    iEndCount = iEndCount - 1
                End If
                iStartCount = iStartCount + 1
                'Debug.Print (aTXT(iStartCount))
         Loop
    iStartCount = 0 'поиск начала раздела с характеристикой
    iStartCount = Find(iStartCount, "ХАРАКТЕРИСТИКИ", aTXT)
    iStartCount = iStartCount + 1
    iEndCount = 0 'поиск конца раздела с характеристикой
    iEndCount = Find(iEndCount, "СПЕЦИАЛЬНАЯ", aTXT)
    iEndCount = iEndCount - 1
         Do While iStartCount < iEndCount
                sData = aTXT(iStartCount)
                sData = Right(sData, Len(sData) - 2) 'удаление одного символа слева
                sData = Left(sData, Len(sData) - 1) 'удаление одного симовола справа
                sData = sDFT + sData
                TrueSeparate (sData)
                iStartCount = iStartCount + 1
         Loop
'____________________________________________________________________________________________________
'конец блока обработки характеристик
'____________________________________________________________________________________________________


iStartCount = 0 'поиск начала раздела с указанием целей
iStartCount = Find(iStartCount, "ДЕЯТЕЛЬНОСТЬ", aTXT)
iStartCount = iStartCount + 5
iEndCount = 0 'поиск конца раздела с указанием
iEndCount = Find(iEndCount, "Команда", aTXT)
iEndCount = iEndCount - 3
iMass1X = 0

XoX = XoX + 1
Loop
End Sub



