Attribute VB_Name = "Module3"
Function TrueSeparate(sStroka As String)
Dim sMassiv(40) As String, sYear As String
Dim iI As Integer, iJ As Integer, iKaretka As Integer, iCounter As Integer
Dim iDemention As Integer, sDate As String

sDate = Left(sStroka, 10)
iDemention = Len(sStroka) - 10 '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������


sMassiv(0) = Left(sStroka, InStr(sStroka, " ") - 1) '���������� ����
iDemention = Len(sStroka) - InStr(sStroka, " ") - 1 '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(1) = Left(sStroka, InStr(sStroka, " ") - 2) 
    If sMassiv(1) <> "�/�" Then  '
        If InStr(sMassiv(1), "-") > 0 Then
            iSep = Len(sMassiv(1)) - InStr(sMassiv(1), "-")
            sMassiv(2) = Right(sMassiv(1), iSep)
            sMassiv(1) = Left(sMassiv(1), InStr(sMassiv(1), "-") - 1)
        Else '���� �� �������� ������, �� ����������� � ������
            sMassiv(2) = sMassiv(1)
            sMassiv(1) = ""
        End If
    End If                          '����� ���������� �������
iDemention = Len(sStroka) - InStr(sStroka, " ") '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(3) = Left(sStroka, InStr(sStroka, "(") - 2) '���������� ��������� ������
iDemention = Len(sStroka) - InStr(sStroka, "(")  '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(4) = Left(sStroka, InStr(sStroka, ")") - 1) '���������� ���� � �������
    If InStr(sMassiv(4), " ") > 0 Then '������ ���� �������
        iSep = Len(sMassiv(4)) - InStr(sMassiv(4), " ")
        sMassiv(5) = Right(sMassiv(4), iSep)
        sMassiv(4) = Left(sMassiv(4), InStr(sMassiv(4), " ") - 1)
        sYear = Right(sDate, 5)
        sMassiv(5) = sMassiv(5) + sYear
    Else
        sMassiv(5) = sDate '���� ���� �� ����� �� ��������� �������� ���� �� ���������� ����
    End If
iDemention = Len(sStroka) - InStr(sStroka, "-") - 1 '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(6) = Left(sStroka, InStr(sStroka, " - ") - 1) '���������� ������
iDemention = Len(sStroka) - InStr(sStroka, " - ") - 2 '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(7) = Left(sStroka, InStr(sStroka, "(") - 2) '���������� ��������� ������
iDemention = Len(sStroka) - InStr(sStroka, "(")  '���������� ������
sStroka = Right(sStroka, iDemention) '���������� ������
sMassiv(8) = Left(sStroka, InStr(sStroka, ")") - 1) '���������� ���� � �������
    If InStr(sMassiv(8), " ") > 0 Then '������ ���� �������
        iSep = Len(sMassiv(8)) - InStr(sMassiv(8), " ")
        sMassiv(9) = Right(sMassiv(8), iSep)
        sMassiv(8) = Left(sMassiv(8), InStr(sMassiv(8), " ") - 1)
        sYear = Right(sDate, 5)
        sMassiv(9) = sMassiv(9) + sYear
    Else
        sMassiv(9) = sDate '���� ���� �� ����� �� ��������� �������� ������� ����
    End If

iI = 1
iJ = 1
Do While Sheets("���").Cells(iI, iJ).Value <> ""
    iI = iI + 1
Loop
Do While iJ <> 11
iCounter = iJ - 1
Sheets("���").Cells(iI, iJ).Value = sMassiv(iCounter)
iJ = iJ + 1
Loop
Sheets("�����").Cells(iI, 13).Value = "�� ������ Scramble"
End Function
Function Find(ByVal iC1 As Integer, sFS1 As String, aMass() As String) As Integer
    Do Until InStr(aMass(iC1), sFS1) <> 0
        iC1 = iC1 + 1
    Loop
    Find = iC1
End Function
Function DelSpace(ByVal sPhrase As String) As String
    sPhrase = Trim(sPhrase) '�������� �������� ����� � ������
        Do While InStr(sPhrase, "  ")
            sPhrase = Replace(sPhrase, "  ", " ") '������ ������� �������� ����������
        Loop
    sPhrase = Right(sPhrase, Len(sPhrase) - 1) '�������� ������ ������� �����
    sPhrase = Left(sPhrase, Len(sPhrase) - 1) '�������� ������ �������� ������
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
'���� ����������� ��������� � ������
'______________________________________________________________________________________________________
  sPut = "D:\�����\����������\1\"
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
        iQuantOfLines = 0   '����������� ���� ������������ ������� ����� � �������� ���������
            Do Until EOF(1)
                Line Input #1, sData
                iQuantOfLines = iQuantOfLines + 1
            Loop
        Close #1
    ReDim aTXT(iQuantOfLines) '�������� ����������� ������� ��� ����������� � ���� ���������
        Open sPathOfTXT For Input As #1
        iQuantOfLines = 0   '����������� ���� ������������ ������� ����� � �������� ���������
            Do Until EOF(1)
                Line Input #1, aTXT(iQuantOfLines)
                iQuantOfLines = iQuantOfLines + 1
            Loop
        Close #1
'______________________________________________________________________________________________________
'����� ����� ����������� ���������
'______________________________________________________________________________________________________


'______________________________________________________________________________________________________
'������ ����� ��������� ����
'______________________________________________________________________________________________________
    
    iFindDate = 1
    iFindDate = Find(iFindDate, "������", aTXT)
    iFindDate = iFindDate + 4
    aTXT(iFindDate) = Trim(aTXT(iFindDate))
    Poopi = aTXT(iFindDate)
    Poopi = Right(Poopi, 11)
    Poopi = Left(Poopi, 6)
    sDFT = Poopi & 2020
'______________________________________________________________________________________________________
'����� ����� ��������� ����
'______________________________________________________________________________________________________


'____________________________________________________________________________________________________
'������ ����� ��������� �������������
'____________________________________________________________________________________________________
    iStartCount = 0 '����� ������ ������� � ���������������
    iStartCount = Find(iStartCount, "������", aTXT)
    iStartCount = iStartCount + 1
    iEndCount = 0 '����� ����� ������� � ���������������
    iEndCount = Find(iEndCount, "�����", aTXT)
    iEndCount = iEndCount - 2
        Do While iStartCount < iEndCount
            iNextString = iStartCount + 1
            
                If Left(aTXT(iNextString), InStr(aTXT(iNextString), " ") - 1) <> "-" Then '���� ��������� ������ ��������
                    aTXT(iStartCount) = aTXT(iStartCount) + aTXT(iNextString)             '������ ����������, �� ��� ������
                        Do While iNextString <> iQuantOfLines                             '�����������, � ������ ���������
                            iCounter = iNextString                                        '�� ���� ������
                            iNextString = iNextString + 1
                            aTXT(iCounter) = aTXT(iNextString)
                        Loop
                    iQuantOfLines = iQuantOfLines - 1
                    iEndCount = iEndCount - 1
                End If
                iStartCount = iStartCount + 1
                'Debug.Print (aTXT(iStartCount))
         Loop
    iStartCount = 0 '����� ������ ������� � ���������������
    iStartCount = Find(iStartCount, "��������������", aTXT)
    iStartCount = iStartCount + 1
    iEndCount = 0 '����� ����� ������� � ���������������
    iEndCount = Find(iEndCount, "�����������", aTXT)
    iEndCount = iEndCount - 1
         Do While iStartCount < iEndCount
                sData = aTXT(iStartCount)
                sData = Right(sData, Len(sData) - 2) '�������� ������ ������� �����
                sData = Left(sData, Len(sData) - 1) '�������� ������ �������� ������
                sData = sDFT + sData
                TrueSeparate (sData)
                iStartCount = iStartCount + 1
         Loop
'____________________________________________________________________________________________________
'����� ����� ��������� �������������
'____________________________________________________________________________________________________


iStartCount = 0 '����� ������ ������� � ��������� �����
iStartCount = Find(iStartCount, "������������", aTXT)
iStartCount = iStartCount + 5
iEndCount = 0 '����� ����� ������� � ���������
iEndCount = Find(iEndCount, "�������", aTXT)
iEndCount = iEndCount - 3
iMass1X = 0

XoX = XoX + 1
Loop
End Sub



