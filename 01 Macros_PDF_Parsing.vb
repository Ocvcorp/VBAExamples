
Public Function FileFolderExists(strFullPath) As Boolean
On Error GoTo EarlyExit
If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
On Error GoTo 0
End Function

Sub InvoiceParse()

Range("A2:K1000").Clear
Range("A2:K1000").HorizontalAlignment = xlCenter
Range("A2:K1000").VerticalAlignment = xlCenter
Range("A2:K1000").WrapText = True



fullPdfPath = Application.GetOpenFilename(",*.pdf")
If FileFolderExists(fullPdfPath) = False Then
    GoTo cancelOpenFile
End If
fullTxtName = Replace(fullPdfPath, ".pdf", ".txt")
strPath = ActiveWorkbook.Path & "\"

strExec = Chr(34) & Chr(34) & strPath & "pdftotext.exe" & Chr(34) & " -layout " & Chr(34) & fullPdfPath & Chr(34) & Chr(34)
Call Shell("cmd /C " & strExec, vbMinimizedNoFocus)
'wait for text file maximimum 5 seconds
Sec = 0
Do Until FileFolderExists(fullTxtName) = True Or Sec > 5
    Application.Wait (Now + TimeValue("0:00:01"))
    Sec = Sec + 1
Loop

'check if text file is there
 blTxt = False
 If FileFolderExists(fullTxtName) = True Then
     blTxt = True
 End If
 
 If blTxt = True Then
         


     ReDim fileStrings(0)
     fsCount = 0
     
     Open fullTxtName For Input As #1
         Do While Not EOF(1)
          Line Input #1, s
          fileStrings(fsCount) = s
          fsCount = fsCount + 1
          ReDim Preserve fileStrings(fsCount)
         Loop
     Close 1
         
    Kill fullTxtName
 End If
    
uBndFS = UBound(fileStrings)

Dim myRegExp As New RegExp
myRegExp.Global = False
myRegExp.IgnoreCase = True
myRegExp.MultiLine = False
'номер инвойса
myRegExp.Pattern = "\d+"
pdfContentString = fileStrings(0)
If myRegExp.test(pdfContentString) Then
    Set matches = myRegExp.Execute(pdfContentString)
End If
invoiceNum = Trim(matches(0).Value)
'дата инвойса
'поиск номера строки с датой
myRegExp.Pattern = "Invoice date"
For i = 1 To uBndFS
    If myRegExp.test(fileStrings(i)) Then
        foundCount = i + 1
        Exit For
    End If
Next i
myRegExp.Pattern = "(\d{1,4}-)\s*(\d{1,2}-)\s*(\d{1,4})$"
pdfContentString = fileStrings(foundCount)
If myRegExp.test(pdfContentString) Then
    Set matches = myRegExp.Execute(pdfContentString)
End If
invoiceDate = Trim(matches(0).Value)
'заполнение таблицы
'поиск стартовой строки
myRegExp.Pattern = "Pos\s*Part\s*no.\s*Name\s*Qty\s*Price"
For i = foundCount To uBndFS
    If myRegExp.test(fileStrings(i)) Then
        foundCount = i + 2
        Exit For
    End If
Next i

'табличные данные
ReDim pos(0)
ReDim partNo(0)
ReDim posInfo(2, 0)
ReDim posQty(0)
ReDim posPriceEach(0)
ReDim posAmount(0)

posCount = 0
str1Complete = False
str2Complete = False
pdfContentString1 = ""
pdfContentString2 = ""
pdfContentString3 = ""
pdfContentString4 = ""
fillOrder1 = -1
fillOrder2 = -1

i = foundCount
'alarmOut = i
PreffixStringAfter = "" 'может случится , что в столбец "Name" в конец некоторой позиции попадет начало следующей строки

For i = foundCount To uBndFS ' And alarmOut <= uBndFS
    myRegExp.Pattern = "^\s{2,8}\d+\s*\S*\s{10,30}\S" 'строка начинается с цифры ^\s{2,8}\d
    If myRegExp.test(fileStrings(i)) Then
        fillOrder1 = 1
    End If
    Select Case fillOrder1
        Case 1
            pdfContentString1 = Trim(fileStrings(i))
            fillOrder1 = 2
        Case 2
            myRegExp.Pattern = "^\s{10,40}\w"
            If myRegExp.test(fileStrings(i)) Then
                pdfContentString2 = Trim(fileStrings(i))
            End If
            fillOrder1 = 3
        Case 3
            myRegExp.Pattern = "^\s{10,40}\w"
            If myRegExp.test(fileStrings(i)) Then
                pdfContentString3 = Trim(fileStrings(i))
            End If
            str1Complete = True
    End Select
    
    myRegExp.Pattern = "^\s{2,8}\d+\s*\S*\s{30,80}\S" 'случай, когда первая строчка не содержит наименования товара
    If myRegExp.test(fileStrings(i)) Then
        fillOrder2 = 1
    End If
    Select Case fillOrder2
        Case 1
            pdfContentString1 = Trim(fileStrings(i))
            fillOrder2 = 2
        Case 2
            myRegExp.Pattern = "^\s{10,40}\w"
            If myRegExp.test(fileStrings(i)) Then
                pdfContentString2 = Trim(fileStrings(i))
            End If
            fillOrder2 = 3
        Case 3
            myRegExp.Pattern = "^\s{10,40}\w"
            If myRegExp.test(fileStrings(i)) Then
                pdfContentString3 = Trim(fileStrings(i))
            End If
            fillOrder2 = 4
        Case 4
            myRegExp.Pattern = "^\s{10,40}\w"
            If myRegExp.test(fileStrings(i)) Then
                pdfContentString4 = Trim(fileStrings(i))
            End If
            str2Complete = True
    End Select
'разбор по шаблону 1
    If str1Complete Then
        ReDim Preserve pos(posCount)
        ReDim Preserve partNo(posCount)
        ReDim Preserve posInfo(2, posCount)
        ReDim Preserve posQty(posCount)
        ReDim Preserve posPriceEach(posCount)
        ReDim Preserve posAmount(posCount)

        'номер позиции
        myRegExp.Pattern = "\s*(\d+)"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            pos(posCount) = Trim(matches(0).Value)
            pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
            InStr(pdfContentString1, pos(posCount)) - Len(pos(posCount)))
        End If
       
        '2. артикул (Part no.)
        'myRegExp.Pattern = "G(?:\d|\w)+(-\s)?(?:\d|\w)+"
        myRegExp.Pattern = "\s*(\S*)?\s*"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            
            partNo(posCount) = Trim(matches(0).Value)
            pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
            InStr(pdfContentString1, partNo(posCount)) - Len(partNo(posCount)))
        End If
        
        '3. наименования (Part no.)
        'возможна ситуация, когда наименование окажется на следующей строке, а цены и кол-во товара на этой строке
        
        myRegExp.Pattern = "\s*(\S.*?\s?\S)\s{5,}"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            
            nameString = Trim(matches(0).Value) & " " & pdfContentString2 & " " & pdfContentString3 ' & " " & Trim(fileStrings(i + 1)) & " " & Trim(fileStrings(i + 2))
            
            myRegExp.Pattern = "(.*)Country of origin:\s*(\w*)\s*\w*\s*weight\s*(\d*\s*.\d*)"
            
            Set matches = myRegExp.Execute(nameString)
            posInfo(0, posCount) = Trim(matches(0).SubMatches.Item(0))
            pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
            InStr(pdfContentString1, posInfo(0, posCount)) - Len(posInfo(0, posCount)))
            posInfo(0, posCount) = Trim(PreffixStringAfter & " " & posInfo(0, posCount))
            
            posInfo(1, posCount) = Trim(matches(0).SubMatches.Item(1))
            posInfo(2, posCount) = Trim(matches(0).SubMatches.Item(2))
            myRegExp.Pattern = "(.*)Kg(.*)"
            Set matches = myRegExp.Execute(posInfo(2, posCount))
            If myRegExp.test(posInfo(2, posCount)) Then
                posInfo(2, posCount) = Trim(matches(0).SubMatches.Item(0))
                PreffixStringAfter = Trim(matches(0).SubMatches.Item(1))
            Else
                PreffixStringAfter = ""
            End If
        End If
        
        '4.
        myRegExp.Pattern = "\s*(\S*\S)\s\S*\s*(\S*\S)\s*(\S*\S)"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)

            
            posQty(posCount) = Trim(matches(0).SubMatches.Item(0))
            posPriceEach(posCount) = Trim(matches(0).SubMatches.Item(1))
            posAmount(posCount) = Trim(matches(0).SubMatches.Item(2))
            
            
        End If
        posCount = posCount + 1
        pdfContentString1 = ""
        fillOrder1 = -1
        str1Complete = False
        'i = i + 2
    End If

'разбор по шаблону 2
    If str2Complete Then
        ReDim Preserve pos(posCount)
        ReDim Preserve partNo(posCount)
        ReDim Preserve posInfo(2, posCount)
        ReDim Preserve posQty(posCount)
        ReDim Preserve posPriceEach(posCount)
        ReDim Preserve posAmount(posCount)
        
        'номер позиции
        myRegExp.Pattern = "\s*(\d+)"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            pos(posCount) = Trim(matches(0).Value)
            pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
            InStr(pdfContentString1, pos(posCount)) - Len(pos(posCount)))
        End If
        
        '2. артикул (Part no.)
        myRegExp.Pattern = "\s*(\S*)?\s*"
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            partNo(posCount) = Trim(matches(0).Value)
            pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
            InStr(pdfContentString1, partNo(posCount)) - Len(partNo(posCount)))
        End If
        
        '3.суммы и количество
        myRegExp.Pattern = "\s*(\S*\S)\s\S*\s*(\S*\S)\s*(\S*\S)"
        pdfContentString1 = Trim(pdfContentString1)
        If myRegExp.test(pdfContentString1) Then
            Set matches = myRegExp.Execute(pdfContentString1)
            posQty(posCount) = Trim(matches(0).SubMatches.Item(0))
            posPriceEach(posCount) = Trim(matches(0).SubMatches.Item(1))
            posAmount(posCount) = Trim(matches(0).SubMatches.Item(2))
        End If
        
        '4.наименование, страна, вес
        nameString = pdfContentString2 & " " & pdfContentString3 & " " & pdfContentString4
        myRegExp.Pattern = "(.*)Country of origin:\s*(\w*)\s*\w*\s*weight\s*(\d*\s*.\d*)"
            
        Set matches = myRegExp.Execute(nameString)
        posInfo(0, posCount) = Trim(matches(0).SubMatches.Item(0))
        pdfContentString1 = Right(pdfContentString1, Len(pdfContentString1) - _
        InStr(pdfContentString1, posInfo(0, posCount)) - Len(posInfo(0, posCount)))
        posInfo(1, posCount) = Trim(matches(0).SubMatches.Item(1))
        posInfo(2, posCount) = Trim(matches(0).SubMatches.Item(2))
        
        posCount = posCount + 1
        pdfContentString1 = ""
        fillOrder2 = -1
        str2Complete = False
        
    End If

Next i



'заполняем таблицу эксель
startRow = 2
For i = 0 To UBound(pos)

    Cells(i + startRow, 1) = invoiceNum
    Cells(i + startRow, 2) = invoiceDate
    Cells(i + startRow, 3) = partNo(i)
    Cells(i + startRow, 4) = posInfo(0, i)
    Cells(i + startRow, 9) = posInfo(1, i)
    Cells(i + startRow, 10) = string2decimal(posInfo(2, i))
    Cells(i + startRow, 5) = string2decimal(posQty(i))
    Cells(i + startRow, 6) = string2decimal(posPriceEach(i))
    Cells(i + startRow, 8) = string2decimal(posAmount(i)) 'cdec

Next i

cancelOpenFile:

End Sub

Function string2decimal(someString)
'функция, преобразующая строку в десятичное число
'убираем из строки все разряды, кроме десятичной дроби
Dim myRegExp As New RegExp
myRegExp.Global = True
myRegExp.Pattern = "\.|,"
If myRegExp.test(someString) Then
    Set matches = myRegExp.Execute(someString)
    If matches.Count > 1 Then
        For i = 0 To matches.Count - 2
            someString = Replace(someString, matches(i), "", 1, 1)
        Next i
    End If
End If
someString = Replace(someString, ".", ",")
string2decimal = CDec(someString)

End Function









