Attribute VB_Name = "ExportKoeffUseFromDatabase"
Sub button_showLumList()
    lumGroupsList.Show
End Sub

Sub button_importCUtables()
cuPathSheet = ActiveSheet.Cells(4, 3)
dirName = ActiveSheet.Cells(5, 3)
sRow = ActiveSheet.Cells(6, 3)

cuImportSheet = ActiveSheet.Cells(7, 3)

importCUtables dirName, cuPathSheet, sRow, cuImportSheet

End Sub

Sub button_makeWordTables()
    makeWordTables
End Sub

Sub test()
MsgBox "it works!"
End Sub

Sub makeWordTables() '��������� �������� ������ � word
'��������� ������ �������
ReDim tableTemp_(2, 0) '������ ������� �������
ReDim wordSheetTables(2, 0): wRC_ = 0 '������ �������� ������
celltext_ = Sheets("����� �������").Cells(2, 1): i = 0
Do While celltext_ <> ""
    ReDim Preserve tableTemp_(2, i)
    tableTemp_(0, i) = celltext_ '����
    tableTemp_(1, i) = Sheets("����� �������").Cells(2 + i, 2) '������ �������
    tableTemp_(2, i) = Sheets("����� �������").Cells(2 + i, 3) '������ �������
    i = i + 1
    celltext_ = Sheets("����� �������").Cells(2 + i, 1)
Loop
tableTemp_ = M_transpose(tableTemp_) '������������� ��� ��������

'��������� ������� ����������� �� ������ ���� ����� ������
sheetTablesCount = 0
generalNamesCount = 0
tT0 = 0
For i = 0 To UBound(tableTemp_)
    If tableTemp_(i, 0) = 1 Then serialNamesCount = serialNamesCount + 1
    If tableTemp_(i, 0) = 2 Then sheetTablesCount = sheetTablesCount + 1
    If sheetTablesCount = 3 Then '��� ������ ���������� 3 ������� �� ���������, ����� �������, ��� ���� � ��� ��������
        ReDim Preserve wordSheetTables(2, wRC_)
        wordSheetTables(0, wRC_) = serialNamesCount + sheetTablesCount * 16 + (sheetTablesCount - 1) * 2 '���-�� ����� ������� � �����
        wordSheetTables(1, wRC_) = tT0 '��������� ������ � ������� �������
        wordSheetTables(2, wRC_) = i '�������� ������ � ������� �������
        sheetTablesCount = 0
        serialNamesCount = 0
        tT0 = i + 1
        wRC_ = wRC_ + 1
    End If
    '������ ��������� ������ �������-������� (����� ����), ���� �� ��������� �� 3 ����
    If sheetTablesCount > 0 And sheetTablesCount < 3 And i = UBound(tableTemp_) Then
        ReDim Preserve wordSheetTables(2, wRC_)
        wordSheetTables(0, wRC_) = serialNamesCount + sheetTablesCount * 16 + (sheetTablesCount - 1) * 2
        wordSheetTables(1, wRC_) = tT0
        wordSheetTables(2, wRC_) = i
    End If
Next i
wordSheetTables = M_transpose(wordSheetTables)

'��������� ������� � ����
'Set objWord = CreateObject("Word.Application")
Set objWord = GetObject(, "Word.Application") '����� ������ "����"
Set objSelection = objWord.Selection '�������� ���
'���������� �� 1-��� ������� Document - ��������������, ��� ��� ������ ��� ����
Set objDoc = objWord.Documents(1) '.Open("D:\����������\�������� ���������� ����������\01 �������� �������\�a��a���u G�llad\14 09 2016 ������� �� �� ����\training2.doc")

objWord.Visible = True '���������� ���

'������� ��� ������� ��������������
For Each oTable In objDoc.Tables
oTable.Delete
Next oTable


'������������ ������ ��� �������
Set objRange = objDoc.Range
'������������� �������� ����� ��� �������/�����
END_OF_STORY = 6
wdpagebreak = 7
 
For iws = 0 To UBound(wordSheetTables)
    '������� ������� � ����
    Nrows = wordSheetTables(iws, 0)
    
    objDoc.Tables.Add objRange, Nrows, 17
    Set objTable = objDoc.Tables(objDoc.Tables.Count) '�������� ��������� ����������� �������
    
    '��������� �� wordSheetTables � tT0 �� tTN
    
    tT0 = wordSheetTables(iws, 1)
    tTN = wordSheetTables(iws, 2)
    iw = 1
    For t = tT0 To tTN
        tKey = tableTemp_(t, 0)
        
        Select Case tKey
            Case 1
                objTable.Cell(iw, 1).Range.Text = tableTemp_(t, 1) '�������� ��������
                '���������� ������ � 1 �� 17
                With objTable
                    Set Rng = .Cell(iw, 1).Range
                    Rng.End = .Cell(iw, 17).Range.End
                    Rng.Cells.Merge
                    .Cell(iw, 1).Range.Bold = 1 '����� ������
                    Rng.ParagraphFormat.Alignment = wdAlignParagraphleft '������������ �����
                End With
                iw = iw + 1
            Case 2
                start_iw = iw
                For colNum = 1 To 2
             
                    If tableTemp_(t, colNum) <> "" Then
                        '�����
                        shift = (colNum - 1) * 9
                        objTable.Cell(iw, 1 + shift - (colNum - 1) * 7).Range.Text = tableTemp_(t, colNum)
                        '���������� 8 �����
                        With objTable
                            Set Rng = .Cell(iw, 1 + shift - (colNum - 1) * 7).Range
                            Rng.End = .Cell(iw, 8 + shift - (colNum - 1) * 7).Range.End
                            Rng.Cells.Merge
                            'Rng.ParagraphFormat.Alignment = 1
                        End With
                        '������������ �����
                        objTable.Cell(iw, 1 + (colNum - 1) * 2).Range.ParagraphFormat.Alignment = 0
                        iw = iw + 1
                        '���
                            '����� �������� � ������� "���������������� ���� ��" �� �������� ������������ �����
                            srch_str = tableTemp_(t, colNum)
                            Set SearchCell = Sheets("���������������� ���� ��").Cells.Find(what:=srch_str, searchformat:=False)
                        '���������� � %
                        wText = Sheets("���������������� ���� ��").Cells(SearchCell.Row, 8) * 100
                        '������
                        objTable.Cell(iw, 1 + shift - (colNum - 1) * 7).Range.Text = "���: " & Format(wText, "#0") & "%"
                        '���������� 8 �����
                        With objTable
                            Set Rng = .Cell(iw, 1 + shift - (colNum - 1) * 7).Range
                            Rng.End = .Cell(iw, 8 + shift - (colNum - 1) * 7).Range.End
                            Rng.Cells.Merge
                            'Rng.ParagraphFormat.Alignment = 1
                        End With
                        '������������ �����
                        objTable.Cell(iw, 1 + (colNum - 1) * 2).Range.ParagraphFormat.Alignment = 0
                        iw = iw + 1
                        '������������ ��������� (����� �� ������������� ������� �� ����� "������ ���� ��")
                        For ri = 1 To 3
                            For rj = 1 To 7
                                objTable.Cell(iw + ri - 1, 1 + rj + shift).Range.Text = Sheets("������ ���� ��").Cells(ri + 1, rj + 9)
                                objTable.Cell(iw + ri - 1, 1 + rj + shift).Shading.BackgroundPatternColor = 12632256 '���� ����
                            Next rj
                        Next ri
                        '������� "rho"

                            objTable.Cell(iw, 1 + shift).Range.Text = ChrW(961) '��������� ��������� ������
                            objTable.Cell(iw, 1 + shift).Range.Bold = 1 '����� ������
                            objTable.Cell(iw, 1 + shift).Range.ParagraphFormat.Alignment = 2 '����������� ������
                            '
                        '������� "i"
                            objTable.Cell(iw + 2, 1 + shift).Range.Text = "i"
                            objTable.Cell(iw + 2, 1 + shift).Range.Bold = 1 '������������ ����������
                        iw = iw + 3
                        '�������� ������������ �������������
                        '����, �� ����� "���������������� ���� ��", ����� ����, � �������� ���������� ������� �� �� ����� "������ ���� ��"
                            srch_str = tableTemp_(t, colNum)
                            Set SearchCell = Sheets("���������������� ���� ��").Cells.Find(what:=srch_str, searchformat:=False)
                        cuExcelRow = (Sheets("���������������� ���� ��").Cells(SearchCell.Row, 1) - 1) * 13 + 3 '��� � �������� ���������� �������� ��
                        KtableMax = Sheets("���������������� ���� ��").Cells(SearchCell.Row, 6) '������������ �� � �������� �������
                        KNormMax = Sheets("���������������� ���� ��").Cells(SearchCell.Row, 7) '���� ��, �� ������� �������� ����������
                        Knorm = KNormMax / KtableMax '����������� �� ������������ �������� ����������� (��������, ����� �� �� ���� ����� 1)
                        For cui = 1 To 11
                            For cuj = 1 To 8
                                If cuj = 1 Then '������� ��������
                                    wText_ = Sheets("������ ���� ��").Cells(cuExcelRow + cui - 1, cuj)
                                    objTable.Cell(iw + cui - 1, cuj + shift).Shading.BackgroundPatternColor = 12632256
                                Else '�������� ��
                                    wText_ = Sheets("������ ���� ��").Cells(cuExcelRow + cui - 1, cuj) * Knorm
                                    wText_ = Format(wText_, "#,##0.00")
                                    objTable.Cell(iw + cui - 1, cuj + shift).Borders.Enable = True
                                End If
                                objTable.Cell(iw + cui - 1, cuj + shift).Range.Text = wText_ '������
                            Next cuj
                        Next cui
                        
                    End If
                    iw = start_iw
                Next colNum
                iw = iw + 18
        End Select
    
    Next t
    '��������� � ���� ����� ��������
    objSelection.endkey END_OF_STORY '�������� ������ � ����� �������
    objSelection.typeparagraph '�������� ������ � ����� �������
    objSelection.insertbreak wdpagebreak '��������� ����� ��������
    Set objRange = objSelection.Range '����������� ������� ����� ��������� �������

Next iws
End Sub



'������ �������� ������ ������������� ������������� �� �� 
Sub importCUtables(dirImportPath, pathsSheetName, startRowPaths, outPutSheetName)

'Copy_fromFolder_toFolder(init_folder, distin_folder)

'��������� ����
'Dim Shablon$, OnlyName$
'Shablon = "*.*": OnlyName = Dir(cuPath_ & Shablon, vbReadOnly + vbHidden + vbSystem)
mainIndex_ = 1 ' �������� ��� ������ ������ ������������

iRow = startRowPaths
fName = Sheets(pathsSheetName).Cells(iRow, 5)
Do While fName <> ""
   '���������������� ����������
   
   '��������� ���� � ��������� ��� ���������� � ���� ������
    Open dirImportPath & fName For Input As #1 '��������� ���� �������� Open() �� ������
    ReDim fileArray_(0): fileArray_(0) = "": stringNum = 0
    Do While Not EOF(1) '���� ���� �� ��������
        ReDim Preserve fileArray_(stringNum)
        Line Input #1, cuString
        fileArray_(stringNum) = cuString
        stringNum = stringNum + 1
    Loop
    Close #1 ' ��������� ����
    
   '��������� ������� � ������ ������ � ������
   '1. ������������ �����������
   Sheets(outPutSheetName).Cells(mainIndex_, 1) = Sheets(pathsSheetName).Cells(iRow, 3)
   '2.������������ ���������
   reflArray_ = Split(fileArray_(1), vbTab)
   For j = 0 To UBound(reflArray_)
        Sheets(outPutSheetName).Cells(1 + mainIndex_, j + 1) = reflArray_(j)
   Next j
   '3.������������ �������������
   For i = 2 To UBound(fileArray_)
        lightIndArray_ = Split(fileArray_(i), vbTab)
        For j = 0 To UBound(lightIndArray_)
            Sheets(outPutSheetName).Cells(i + mainIndex_, j + 1) = lightIndArray_(j)
        Next j
        
   Next i
   mainIndex_ = mainIndex_ + i


   '�����--------���������������� ����������
   iRow = iRow + 1
   fName = Sheets(pathsSheetName).Cells(iRow, 5)
Loop

End Sub


Sub stringFilters()
fType_ = "DblSpace"
Select Case fType_
Case "Space"
    For i = 210 To 224
        string_ = Trim(ActiveSheet.Cells(i, 4))
        sPos_ = InStr(string_, " ")
        string_ = Left(string_, sPos_ - 1)
        ActiveSheet.Cells(i, 4) = string_
    Next i
Case "DblSpace"
    For i = 181 To 209
        string_ = Trim(ActiveSheet.Cells(i, 4))
        sPos_ = InStr(string_, " ")
        string_ = Left(string_, sPos_ - 1) & Right(string_, Len(string_) - sPos_)
        sPos_ = InStr(string_, " ")
        string_ = Left(string_, sPos_ - 1)
        ActiveSheet.Cells(i, 4) = string_
    Next i
End Select
End Sub


Function M_transpose(fMtrx()) As Variant()
'�������, ������� ���������� ����������������� �������
'�� ���������� ��������� �������
'{
'Dim rowLastNo_, columnLastNo_ As Integer
rowLastNo_ = UBound(fMtrx, 1): columnLastNo_ = UBound(fMtrx, 2)
ReDim bufferMtrx_(columnLastNo_, rowLastNo_)
Dim i, j As Integer
For i = 0 To rowLastNo_
    For j = 0 To columnLastNo_
        bufferMtrx_(j, i) = fMtrx(i, j)
    Next j
Next i
M_transpose = bufferMtrx_
End Function
'}

Function getCellValbyNumber(fSheetName, fRowNum, fColNum)
'�������, ������� �������� �������� �� ������ �� ������ ������� � ���� � ������������ �����
getCellValbyNumber = Sheets(fSheetName).Cells(fRowNum, fColNum)

End Function
