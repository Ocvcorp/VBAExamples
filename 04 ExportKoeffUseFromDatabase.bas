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

Sub makeWordTables() 'процедура переноса таблиц в word
'вычисляем скелет таблицы
ReDim tableTemp_(2, 0) 'массив скелета таблицы
ReDim wordSheetTables(2, 0): wRC_ = 0 'массив размеров таблиц
celltext_ = Sheets("Схема таблицы").Cells(2, 1): i = 0
Do While celltext_ <> ""
    ReDim Preserve tableTemp_(2, i)
    tableTemp_(0, i) = celltext_ 'ключ
    tableTemp_(1, i) = Sheets("Схема таблицы").Cells(2 + i, 2) 'первый столбец
    tableTemp_(2, i) = Sheets("Схема таблицы").Cells(2 + i, 3) 'второй столбец
    i = i + 1
    celltext_ = Sheets("Схема таблицы").Cells(2 + i, 1)
Loop
tableTemp_ = M_transpose(tableTemp_) 'транспонируем для удобства

'вычисляем размеры вставляемых на каждый лист ворда таблиц
sheetTablesCount = 0
generalNamesCount = 0
tT0 = 0
For i = 0 To UBound(tableTemp_)
    If tableTemp_(i, 0) = 1 Then serialNamesCount = serialNamesCount + 1
    If tableTemp_(i, 0) = 2 Then sheetTablesCount = sheetTablesCount + 1
    If sheetTablesCount = 3 Then 'как только набирается 3 таблицы по вертикали, можно считать, что лист у нас заполнен
        ReDim Preserve wordSheetTables(2, wRC_)
        wordSheetTables(0, wRC_) = serialNamesCount + sheetTablesCount * 16 + (sheetTablesCount - 1) * 2 'кол-во строк таблицы в ворде
        wordSheetTables(1, wRC_) = tT0 'начальный индекс в таблице скелета
        wordSheetTables(2, wRC_) = i 'конечный индекс в таблице скелета
        sheetTablesCount = 0
        serialNamesCount = 0
        tT0 = i + 1
        wRC_ = wRC_ + 1
    End If
    'случай последней группы таблицы-скелета (листа Ворд), если по вертикали не 3 ряда
    If sheetTablesCount > 0 And sheetTablesCount < 3 And i = UBound(tableTemp_) Then
        ReDim Preserve wordSheetTables(2, wRC_)
        wordSheetTables(0, wRC_) = serialNamesCount + sheetTablesCount * 16 + (sheetTablesCount - 1) * 2
        wordSheetTables(1, wRC_) = tT0
        wordSheetTables(2, wRC_) = i
    End If
Next i
wordSheetTables = M_transpose(wordSheetTables)

'заполняем таблицы в ворд
'Set objWord = CreateObject("Word.Application")
Set objWord = GetObject(, "Word.Application") 'берем объект "ворд"
Set objSelection = objWord.Selection 'выделяем его
'добираемся до 1-ого объекта Document - предполагается, что это нужный нам файл
Set objDoc = objWord.Documents(1) '.Open("D:\Информация\Каталоги Справочная информация\01 Световые приборы\кaттaлогu Gаllad\14 09 2016 экспорт КИ из БДСП\training2.doc")

objWord.Visible = True 'показываем его

'удаляем все таблицы предварительно
For Each oTable In objDoc.Tables
oTable.Delete
Next oTable


'устанвливаем начало для таблицы
Set objRange = objDoc.Range
'устанавливаем конечную точку для таблицы/листа
END_OF_STORY = 6
wdpagebreak = 7
 
For iws = 0 To UBound(wordSheetTables)
    'создаем таблицу в ворд
    Nrows = wordSheetTables(iws, 0)
    
    objDoc.Tables.Add objRange, Nrows, 17
    Set objTable = objDoc.Tables(objDoc.Tables.Count) 'выбираем последнюю добавленную таблицу
    
    'заполняем по wordSheetTables с tT0 по tTN
    
    tT0 = wordSheetTables(iws, 1)
    tTN = wordSheetTables(iws, 2)
    iw = 1
    For t = tT0 To tTN
        tKey = tableTemp_(t, 0)
        
        Select Case tKey
            Case 1
                objTable.Cell(iw, 1).Range.Text = tableTemp_(t, 1) 'вставили название
                'объединили ячейки с 1 по 17
                With objTable
                    Set Rng = .Cell(iw, 1).Range
                    Rng.End = .Cell(iw, 17).Range.End
                    Rng.Cells.Merge
                    .Cell(iw, 1).Range.Bold = 1 'шрифт жирным
                    Rng.ParagraphFormat.Alignment = wdAlignParagraphleft 'выравнивание влево
                End With
                iw = iw + 1
            Case 2
                start_iw = iw
                For colNum = 1 To 2
             
                    If tableTemp_(t, colNum) <> "" Then
                        'серия
                        shift = (colNum - 1) * 9
                        objTable.Cell(iw, 1 + shift - (colNum - 1) * 7).Range.Text = tableTemp_(t, colNum)
                        'объединяем 8 ячеек
                        With objTable
                            Set Rng = .Cell(iw, 1 + shift - (colNum - 1) * 7).Range
                            Rng.End = .Cell(iw, 8 + shift - (colNum - 1) * 7).Range.End
                            Rng.Cells.Merge
                            'Rng.ParagraphFormat.Alignment = 1
                        End With
                        'выравнивание влево
                        objTable.Cell(iw, 1 + (colNum - 1) * 2).Range.ParagraphFormat.Alignment = 0
                        iw = iw + 1
                        'КПД
                            'поиск значения в таблице "Оптимизированная табл КИ" по краткому наименованию серии
                            srch_str = tableTemp_(t, colNum)
                            Set SearchCell = Sheets("Оптимизированная табл КИ").Cells.Find(what:=srch_str, searchformat:=False)
                        'приведение к %
                        wText = Sheets("Оптимизированная табл КИ").Cells(SearchCell.Row, 8) * 100
                        'запись
                        objTable.Cell(iw, 1 + shift - (colNum - 1) * 7).Range.Text = "КПД: " & Format(wText, "#0") & "%"
                        'объединяем 8 ячеек
                        With objTable
                            Set Rng = .Cell(iw, 1 + shift - (colNum - 1) * 7).Range
                            Rng.End = .Cell(iw, 8 + shift - (colNum - 1) * 7).Range.End
                            Rng.Cells.Merge
                            'Rng.ParagraphFormat.Alignment = 1
                        End With
                        'выравнивание влево
                        objTable.Cell(iw, 1 + (colNum - 1) * 2).Range.ParagraphFormat.Alignment = 0
                        iw = iw + 1
                        'коэффициенты отражения (берем из заготовленной таблицы на листе "Полная табл КИ")
                        For ri = 1 To 3
                            For rj = 1 To 7
                                objTable.Cell(iw + ri - 1, 1 + rj + shift).Range.Text = Sheets("Полная табл КИ").Cells(ri + 1, rj + 9)
                                objTable.Cell(iw + ri - 1, 1 + rj + shift).Shading.BackgroundPatternColor = 12632256 'цвет фона
                            Next rj
                        Next ri
                        'подпись "rho"

                            objTable.Cell(iw, 1 + shift).Range.Text = ChrW(961) 'вставляем греческий символ
                            objTable.Cell(iw, 1 + shift).Range.Bold = 1 'шрифт жирным
                            objTable.Cell(iw, 1 + shift).Range.ParagraphFormat.Alignment = 2 'выравниваем вправо
                            '
                        'подпись "i"
                            objTable.Cell(iw + 2, 1 + shift).Range.Text = "i"
                            objTable.Cell(iw + 2, 1 + shift).Range.Bold = 1 'выравнивание посередине
                        iw = iw + 3
                        'значения коэффицентов использования
                        'ищем, по листу "Оптимизированная табл КИ", номер ряда, с которого начинается таблица КИ на листе "Полная табл КИ"
                            srch_str = tableTemp_(t, colNum)
                            Set SearchCell = Sheets("Оптимизированная табл КИ").Cells.Find(what:=srch_str, searchformat:=False)
                        cuExcelRow = (Sheets("Оптимизированная табл КИ").Cells(SearchCell.Row, 1) - 1) * 13 + 3 'ряд с которого начинаются значения КИ
                        KtableMax = Sheets("Оптимизированная табл КИ").Cells(SearchCell.Row, 6) 'максимальный КИ в исходной таблице
                        KNormMax = Sheets("Оптимизированная табл КИ").Cells(SearchCell.Row, 7) 'макс КИ, на который задается нормировка
                        Knorm = KNormMax / KtableMax 'нормирующий на определенное значение коэффициент (поправка, чтобы КИ не были более 1)
                        For cui = 1 To 11
                            For cuj = 1 To 8
                                If cuj = 1 Then 'столбец индексов
                                    wText_ = Sheets("Полная табл КИ").Cells(cuExcelRow + cui - 1, cuj)
                                    objTable.Cell(iw + cui - 1, cuj + shift).Shading.BackgroundPatternColor = 12632256
                                Else 'значения КИ
                                    wText_ = Sheets("Полная табл КИ").Cells(cuExcelRow + cui - 1, cuj) * Knorm
                                    wText_ = Format(wText_, "#,##0.00")
                                    objTable.Cell(iw + cui - 1, cuj + shift).Borders.Enable = True
                                End If
                                objTable.Cell(iw + cui - 1, cuj + shift).Range.Text = wText_ 'запись
                            Next cuj
                        Next cui
                        
                    End If
                    iw = start_iw
                Next colNum
                iw = iw + 18
        End Select
    
    Next t
    'вставляем в ворд конец страницы
    objSelection.endkey END_OF_STORY 'помещаем курсор в конец таблицы
    objSelection.typeparagraph 'помещаем курсор в конец таблицы
    objSelection.insertbreak wdpagebreak 'добавляем новую страницу
    Set objRange = objSelection.Range 'присваиваем области новое положение курсора

Next iws
End Sub



'Модуль экспорта таблиц коэффициентов использования из БД 
Sub importCUtables(dirImportPath, pathsSheetName, startRowPaths, outPutSheetName)

'Copy_fromFolder_toFolder(init_folder, distin_folder)

'считываем файл
'Dim Shablon$, OnlyName$
'Shablon = "*.*": OnlyName = Dir(cuPath_ & Shablon, vbReadOnly + vbHidden + vbSystem)
mainIndex_ = 1 ' смещение для таблиц разных светильников

iRow = startRowPaths
fName = Sheets(pathsSheetName).Cells(iRow, 5)
Do While fName <> ""
   'непосредственное считывание
   
   'Открываем файл и считываем всю информацию в одну строку
    Open dirImportPath & fName For Input As #1 'Открываем файл функцией Open() на чтение
    ReDim fileArray_(0): fileArray_(0) = "": stringNum = 0
    Do While Not EOF(1) 'пока файл не кончился
        ReDim Preserve fileArray_(stringNum)
        Line Input #1, cuString
        fileArray_(stringNum) = cuString
        stringNum = stringNum + 1
    Loop
    Close #1 ' Закрываем файл
    
   'обработка массива и запись данных в эксель
   '1. наименование светильника
   Sheets(outPutSheetName).Cells(mainIndex_, 1) = Sheets(pathsSheetName).Cells(iRow, 3)
   '2.коэффициенты отражения
   reflArray_ = Split(fileArray_(1), vbTab)
   For j = 0 To UBound(reflArray_)
        Sheets(outPutSheetName).Cells(1 + mainIndex_, j + 1) = reflArray_(j)
   Next j
   '3.коэффициенты использования
   For i = 2 To UBound(fileArray_)
        lightIndArray_ = Split(fileArray_(i), vbTab)
        For j = 0 To UBound(lightIndArray_)
            Sheets(outPutSheetName).Cells(i + mainIndex_, j + 1) = lightIndArray_(j)
        Next j
        
   Next i
   mainIndex_ = mainIndex_ + i


   'конец--------непосредственное считывание
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
'функция, которая возвращает транспонированную матрицу
'не использует сторонние функции
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
'функция, которая получает значение из ячейки по номеру столбца и ряда и наименованию листа
getCellValbyNumber = Sheets(fSheetName).Cells(fRowNum, fColNum)

End Function
