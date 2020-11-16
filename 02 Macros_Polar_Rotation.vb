Sub main()

'алгоритм поворота

'1. переводим в декартовы координаты полярный массив КСС
'2. поворачиваем каждый вектор на alpha, beta; для alpha от 0 до 180 в соответствующей четверь-плоскости C получается одна линия
'2.1 при этом, начиная со 2ой линии (следующее значение C) создаем массив треугольников поверхности - не забываем в конце про "склейку" С355 и С0 - для это надо, чтобы в исходной КСС была плоскость С360=С0
'3. вычисляем пересечение повернутой поверхности КСС со сферой углов gamma,C в МСК
'3.1 для этого создаем единичный вектор gamma,C
'3.2 перебираем все треугольники, на которые мы разбилии КСС и ищем точку пересечения вектора и треугольника, которая лежит внутри треугольника (или его на сторонах)
'3.3 модуль вектора в этой точке и есть искомое значение КСС
'
'П.С. у алгоритма есть минусы, связанные с неточностью расчетов - надо дорабатывать в будущем

'также абсолютно без объявления типов у меня, что очень неправильно

'------------------реализация программы
'углы поворота оптической оси светильника:
'alpha - вокруг оси Ох в направлении y=1 от z=-1, если смотреть с конца Ox
'beta - округ Oz в направлении y=-1 от x=1, если смотреть с конца Oz

alpha = 0: beta = 30 ' углы поворота оптической оси светильника

dC = 45: dGamma = 15 'шаг полярных углов
NC = 360 / dC + 1: NGamma = 180 / dGamma + 1

ReDim buff_Arr_before(NGamma - 1, 2), buff_Arr_current(NGamma - 1, 2) 'массивы линий gamma для каждого C, чтобы разбить поверхность на треугольники
ReDim empty_arr(0, 2, 2) 'начальный нулевой массив, чтобы не объявлять trind_rotated_KSS и впоследствии реализовать функцию добавления к нему строк - такой кривой способ у меня
trind_rotated_KSS = empty_arr
start_index = 0


startCellRow = 79: startCellColumn = 2
For ii = 0 To NC - 1
    C_LC = dC * ii
    For j = 0 To NGamma - 1
        gamma_LC = dGamma * j
        I_LC = Sheets(2).Cells(startCellRow + ii, startCellColumn + j) 'ввод значений в программу из Excel
        'перевод (I_LC, gamma_LC, C_LC) в ДСК
        Ixyz = polarVec_2_cartesianVec(I_LC, gamma_LC, C_LC)
        Ixyz(0) = Round(Ixyz(0), 0): Ixyz(1) = Round(Ixyz(1), 0): Ixyz(2) = Round(Ixyz(2), 0) 'округляю из-за проблем с точностью, которые надо решать
        
        'поворот вектора на углы alpha, beta
        Ixyz_ab = ldc_rotate(Ixyz, alpha, beta)
        Ixyz_ab(0) = Round(Ixyz_ab(0), 0): Ixyz_ab(1) = Round(Ixyz_ab(1), 0): Ixyz_ab(2) = Round(Ixyz_ab(2), 0)
        
        'формируем линию КСС в каждой четверть-плоскости "С"
        buff_Arr_current(j, 0) = Ixyz_ab(0)
        buff_Arr_current(j, 1) = Ixyz_ab(1)
        buff_Arr_current(j, 2) = Ixyz_ab(2)
        
    Next j
    
    'из двух массивов заполняю массив треугольников
    If ii <> 0 Then 'начиная с ii=1, когда будет уже 2 линии: buff_Arr_before и buff_Arr_current, ведь только из 2х линий можно сделать поверхность (треугольниками)
        n_returned_Arr = triangles_from_2curves(buff_Arr_before, buff_Arr_current)
        trind_rotated_KSS = add_3D_Array(trind_rotated_KSS, n_returned_Arr)
    End If
    
    'по окончании "треуангуляции" записывам линию для текущего угла "С" в стутус предыдущей и переходим к следующему углу "С"
    For ttt = 0 To UBound(buff_Arr_current)
        buff_Arr_before(ttt, 0) = buff_Arr_current(ttt, 0)
        buff_Arr_before(ttt, 1) = buff_Arr_current(ttt, 1)
        buff_Arr_before(ttt, 2) = buff_Arr_current(ttt, 2)
    Next ttt

Next ii

trind_rotated_KSS = delete_firstIndex_3D_array(trind_rotated_KSS, 0) 'удаляем первый пустой ряд

'выводим получившиеся данные в формат таблицы Excel, который потом экспортируем в автокад для визуальной отладки
For ii = 0 To UBound(trind_rotated_KSS)
    For j = 0 To 2
        For k = 0 To 2
            Sheets(2).Cells(90 + ii, 9 + 3 * j + k) = trind_rotated_KSS(ii, j, k) 'Round(, 0)
        Next k
    Next j
Next ii

'определяем интерполяцией значения  повернутой КСС в заданных узлах полярной сетки МСК
ReDim I_MSK(NC - 2, NGamma - 1)
For ii = 0 To NC - 2
    C_MSK = ii * dC
    For j = 0 To NGamma - 1
        gamma_MSK = j * dGamma
        I_MSK(ii, j) = find_KSS(trind_rotated_KSS, gamma_MSK, C_MSK) 'функция перевода
    Next j
Next ii

'выводим финальный результат в эксель
For ii = 0 To NC - 2
    For j = 0 To NGamma - 1
        Sheets(3).Cells(2 + ii, 1) = ii * dC
        Sheets(3).Cells(1, startCellColumn + j) = j * dGamma
        Sheets(3).Cells(2 + ii, startCellColumn + j) = I_MSK(ii, j)
    Next j
Next ii

End Sub

Function polarVec_2_cartesianVec(I, gamma, C)
'функция, которая переводит полярные координаты в декартовы из следующих принятых положений:
'угол gamma отсчитывается от оси oZ=-1 в правой тройке ZYX
'угол С отсчитывается от оси oY=1 в правой тройке XYZ
Dim Ixyz(2)
Ixyz(0) = -I * Sin(rad(gamma)) * Sin(rad(C))
Ixyz(1) = I * Sin(rad(gamma)) * Cos(rad(C))
Ixyz(2) = -I * Cos(rad(gamma))
polarVec_2_cartesianVec = Ixyz
End Function

Function ldc_rotate(vec, alpha, beta)
'
'функция поворота поворота оптической оси светильника:
'alpha - вокруг оси Ох в направлении y=1 от z=-1, если смотреть с конца Ox
'beta - округ Oz в направлении y=-1 от x=1, если смотреть с конца Oz
'матрицы поворота, положительный угол соответствует повороту против часовой стрелки в правой системе координат
Dim Mx(2, 2), Mz(2, 2)

Mx(0, 0) = 1: Mx(0, 1) = 0:               Mx(0, 2) = 0
Mx(1, 0) = 0: Mx(1, 1) = Cos(rad(alpha)): Mx(1, 2) = -Sin(rad(alpha))
Mx(2, 0) = 0: Mx(2, 1) = Sin(rad(alpha)): Mx(2, 2) = Cos(rad(alpha))

Mz(0, 0) = Cos(rad(beta)): Mz(0, 1) = -Sin(rad(beta)):  Mz(0, 2) = 0
Mz(1, 0) = Sin(rad(beta)): Mz(1, 1) = Cos(rad(beta)):   Mz(1, 2) = 0
Mz(2, 0) = 0:              Mz(2, 1) = 0:                Mz(2, 2) = 1

'поворачиваем вектор, последовательно умножая на соответствующие матрицы

ansVec1 = mtrx3x3_to_vec3_Product(Mx, vec)
ansVec2 = mtrx3x3_to_vec3_Product(Mz, ansVec1)

ldc_rotate = ansVec2

End Function

Function triangles_from_2curves(arr_before, arr_current)
'функция, которая соединяет последовательные точки 2х кривых в треугольники
'основные положения:
'1. как минимум в начале и конце кривые имеют общие точки (как максимум и в середине)
'2. точки соединяются последовательно- т.е., например: 1ая кривая т.1+т.2 2кривая + т.1 и 1ая кривая т.2 2кривая + т.1+т.2
'3. необходимо отслеживать места, где грани совпадают - у начала кривых, в конце, в других местах
'признак: тройки точек совпадают (т.е. грани совпадают)
'в этом случае надо выделить треугольную грань (т.е. ту грань, у которой точки не сопадают),
'т.к. другая грань выродиться в отрезок

Dim pnt11(2), pnt12(2), pnt21(2), pnt22(2) 'точки с двух кривых
ReDim ans(2, 2, 0) 'перевернутый массив формата координата-точка-грань
'переворачиваем массив (т.е. 1ый и 3ий индекс меняем местами) из-за того, что надо массив динамический и увеличивать у него в VBA можно только последний индекс

U_ii = UBound(arr_before)
n_tri = -1 'счетчик выходного массива граней

For ii = 0 To U_ii - 1

pnt11(0) = arr_before(ii, 0): pnt11(1) = arr_before(ii, 1): pnt11(2) = arr_before(ii, 2)
pnt12(0) = arr_before(ii + 1, 0): pnt12(1) = arr_before(ii + 1, 1): pnt12(2) = arr_before(ii + 1, 2)
pnt21(0) = arr_current(ii, 0): pnt21(1) = arr_current(ii, 1): pnt21(2) = arr_current(ii, 2)
pnt22(0) = arr_current(ii + 1, 0): pnt22(1) = arr_current(ii + 1, 1): pnt22(2) = arr_current(ii + 1, 2)

face_bufer = trifaces_from_4_points(pnt11, pnt12, pnt21, pnt22) 'функция, которая из 4х точек делает создает или одну или две треугольных поверхности(грани)

    If UBound(face_bufer) = 0 Then 'если функция выдает массив только с одной строкой, значит одна грань стянулась в отрезок (это чаще вначале кривых)
    'случай одной грани
        n_tri = n_tri + 1
        ReDim Preserve ans(2, 2, n_tri)
        '1ая точка:
        ans(0, 0, n_tri) = face_bufer(0, 0, 0): ans(1, 0, n_tri) = face_bufer(0, 0, 1): ans(2, 0, n_tri) = face_bufer(0, 0, 2)
        '2ая точка:
        ans(0, 1, n_tri) = face_bufer(0, 1, 0): ans(1, 1, n_tri) = face_bufer(0, 1, 1): ans(2, 1, n_tri) = face_bufer(0, 1, 2)
        '3я точка:
        ans(0, 2, n_tri) = face_bufer(0, 2, 0): ans(1, 2, n_tri) = face_bufer(0, 2, 1): ans(2, 2, n_tri) = face_bufer(0, 2, 2)
     End If
     If UBound(face_bufer) = 1 Then
        'случай двух граней
        n_tri = n_tri + 2
        ReDim Preserve ans(2, 2, n_tri)
        '1ая грань
            '1ая точка:
            ans(0, 0, n_tri - 1) = face_bufer(0, 0, 0): ans(1, 0, n_tri - 1) = face_bufer(0, 0, 1): ans(2, 0, n_tri - 1) = face_bufer(0, 0, 2)
            '2ая точка:
            ans(0, 1, n_tri - 1) = face_bufer(0, 1, 0): ans(1, 1, n_tri - 1) = face_bufer(0, 1, 1): ans(2, 1, n_tri - 1) = face_bufer(0, 1, 2)
            '3я точка:
            ans(0, 2, n_tri - 1) = face_bufer(0, 2, 0): ans(1, 2, n_tri - 1) = face_bufer(0, 2, 1): ans(2, 2, n_tri - 1) = face_bufer(0, 2, 2)
         '2ая грань
            '1ая точка:
            ans(0, 0, n_tri) = face_bufer(1, 0, 0): ans(1, 0, n_tri) = face_bufer(1, 0, 1): ans(2, 0, n_tri) = face_bufer(1, 0, 2)
            '2ая точка:
            ans(0, 1, n_tri) = face_bufer(1, 1, 0): ans(1, 1, n_tri) = face_bufer(1, 1, 1): ans(2, 1, n_tri) = face_bufer(1, 1, 2)
            '3я точка:
            ans(0, 2, n_tri) = face_bufer(1, 2, 0): ans(1, 2, n_tri) = face_bufer(1, 2, 1): ans(2, 2, n_tri) = face_bufer(1, 2, 2)
     End If

Next ii

triangles_from_2curves = rev_3D_array(ans) 'перестравиваем массив в изначальный формат (т.е. меняем местами элементы с 1ым и 3им индексами)

End Function

Function trifaces_from_4_points(point11, point12, point21, point22)
'функция, которая создает либо одну, либо 2 треугольных грани из 4х точек
'функция охватывает не все комбинации:
'предполагается, что в грани могут соединятся только две тройки точек
'1ая: point11-point12-point21
'2ая: point12-point21-point22
'если совпадают пары point11-point21, point12-point22
'то на выходе получается одна грань, а именно:
'1) если point11-point21, то point12-point21-point22
'2) если point12-point22, то point11-point12-point21

ReDim ans(0, 2, 2)
'случай, когда две грани:
'*  point11-point12-point21
'** point12-point21-point22
If two_points_NOT_overlap(point11, point21) And two_points_NOT_overlap(point12, point22) Then
    ReDim ans(1, 2, 2)
    '1ая грань
        '1ая точка:
        ans(0, 0, 0) = point11(0): ans(0, 0, 1) = point11(1): ans(0, 0, 2) = point11(2)
        '2ая точка:
        ans(0, 1, 0) = point12(0): ans(0, 1, 1) = point12(1): ans(0, 1, 2) = point12(2)
        '3я точка:
        ans(0, 2, 0) = point21(0): ans(0, 2, 1) = point21(1): ans(0, 2, 2) = point21(2)
     '2ая грань
        '1ая точка:
        ans(1, 0, 0) = point12(0): ans(1, 0, 1) = point12(1): ans(1, 0, 2) = point12(2)
        '2ая точка:
        ans(1, 1, 0) = point21(0): ans(1, 1, 1) = point21(1): ans(1, 1, 2) = point21(2)
        '3я точка:
        ans(1, 2, 0) = point22(0): ans(1, 2, 1) = point22(1): ans(1, 2, 2) = point22(2)
End If
'случай, когда одна грань point12-point21-point22
If two_points_NOT_overlap(point11, point21) = False Then
    '1ая точка:
    ans(0, 0, 0) = point12(0): ans(0, 0, 1) = point12(1): ans(0, 0, 2) = point12(2)
    '2ая точка:
    ans(0, 1, 0) = point21(0): ans(0, 1, 1) = point21(1): ans(0, 1, 2) = point21(2)
    '3я точка:
    ans(0, 2, 0) = point22(0): ans(0, 2, 1) = point22(1): ans(0, 2, 2) = point22(2)
End If
'случай, когда одна грань point11-point12-point21
If two_points_NOT_overlap(point12, point22) = False Then
    '1ая точка:
    ans(0, 0, 0) = point11(0): ans(0, 0, 1) = point11(1): ans(0, 0, 2) = point11(2)
    '2ая точка:
    ans(0, 1, 0) = point12(0): ans(0, 1, 1) = point12(1): ans(0, 1, 2) = point12(2)
    '3я точка:
    ans(0, 2, 0) = point21(0): ans(0, 2, 1) = point21(1): ans(0, 2, 2) = point21(2)
End If

trifaces_from_4_points = ans

End Function

Function find_KSS(triangled_Cartesian_KSS, gamma, C)
'функция поиска значения КСС для заданных углов
'1.ищется точка пересечения прямой, заданной (gamma, C) с каждым треугольником в triangled_Cartesian_KSS
'2.если точка пересечения попадает внутрь треугольника, то вычисляется модуль соответствующего вектора - это есть значение КСС

Dim cross_point(2)
ans = 0
ii = 0 'счетчик по массиву граней
Do While ii <= UBound(triangled_Cartesian_KSS, 1) And ans <= 0
    'координаты треугольника
    P1 = Array(triangled_Cartesian_KSS(ii, 0, 0), triangled_Cartesian_KSS(ii, 0, 1), triangled_Cartesian_KSS(ii, 0, 2))
    P2 = Array(triangled_Cartesian_KSS(ii, 1, 0), triangled_Cartesian_KSS(ii, 1, 1), triangled_Cartesian_KSS(ii, 1, 2))
    P3 = Array(triangled_Cartesian_KSS(ii, 2, 0), triangled_Cartesian_KSS(ii, 2, 1), triangled_Cartesian_KSS(ii, 2, 2))
    'плоскость треугольника:
    'нормаль
    BA = Array(P2(0) - P1(0), P2(1) - P1(1), P2(2) - P1(2))
    CA = Array(P3(0) - P1(0), P3(1) - P1(1), P3(2) - P1(2))
    plane_normal = VecProduct(CA, BA)
    abs_plane_normal = VecMod(plane_normal)
    If abs_plane_normal = 0 Then
        ans = -1 'случай, когда плоскость стянулась в точку
    Else
        plane_normal(0) = plane_normal(0) / abs_plane_normal: plane_normal(1) = plane_normal(1) / abs_plane_normal: plane_normal(2) = plane_normal(2) / abs_plane_normal
        'луч
        nRay = polarVec_2_cartesianVec(1, gamma, C)
        buff_ans = find_Crosspoint_of_Ray_and_Plane(nRay, plane_normal, P1)
        
        If buff_ans(3) = 0 Then
            cross_point(0) = buff_ans(0): cross_point(1) = buff_ans(1): cross_point(2) = buff_ans(2)
            If point_is_inOrigin(cross_point, 0.0001) = False Then
                'определяем другую точку
                If point_is_inOrigin(P1, 0.0001) Then
                    other_point = P2
                Else
                    other_point = P1
                End If
                If codirectional_of_vectors(nRay, other_point) Then 'если направлены в одну сторону векторы (crosspoint-(0,0,0)); (other_point-(0,0,0))
                                                                            'т.к. нам надо именно пересечение с гранью луча, а не прямой
                    If point_inside_triangle(P1, P2, P3, cross_point) Then 'если точка пересечения внутри грани
                        ans = VecMod(cross_point)
                    End If
                End If
            End If
        Else
            ans = -2 '!!!случай, когда вектор параллелен плоскости - надо уточнить, нужен он или нет
        End If
    End If
ii = ii + 1
Loop
find_KSS = ans
End Function



'МАТЕМАТИЧЕСКИЕ ФУНКЦИИ
'МАТРИЦЫ
Function mtrx3x3_to_vec3_Product(mtrx3x3, vec3)
'умножение матрицы 3х3 на вектор-столбец
Dim ans(2)
ans(0) = mtrx3x3(0, 0) * vec3(0) + mtrx3x3(0, 1) * vec3(1) + mtrx3x3(0, 2) * vec3(2)
ans(1) = mtrx3x3(1, 0) * vec3(0) + mtrx3x3(1, 1) * vec3(1) + mtrx3x3(1, 2) * vec3(2)
ans(2) = mtrx3x3(2, 0) * vec3(0) + mtrx3x3(2, 1) * vec3(1) + mtrx3x3(2, 2) * vec3(2)
mtrx3x3_to_vec3_Product = ans

End Function


Function add_3D_Array(enlarging_Array, adding_Array)
'функция, увеличивающая трехмерный массив
'кол-во 2ого и 3его измерений остается прежним
'увеличение идет только по 1ому измерению
U_enA = UBound(enlarging_Array, 1): L1_enA = UBound(enlarging_Array, 2): L2_enA = UBound(enlarging_Array, 3)
U_adA = UBound(adding_Array, 1)

ReDim ans(U_enA + U_adA + 1, L1_enA, L2_enA)

For k = 0 To L2_enA
    For j = 0 To L1_enA
        For ii = 0 To U_enA + U_adA + 1
            If ii <= U_enA Then
                ans(ii, j, k) = enlarging_Array(ii, j, k)
            Else
                ans(ii, j, k) = adding_Array(ii - U_enA - 1, j, k)
            End If
        Next ii
    Next j
Next k

add_3D_Array = ans

End Function

Function rev_3D_array(input_3D_array)
'функция преобразущая измерения трехмерного массива в обратном порядке, т.е. 1,2,3 в 3,2,1
U_input_1 = UBound(input_3D_array, 1)
U_input_2 = UBound(input_3D_array, 2)
U_input_3 = UBound(input_3D_array, 3)

ReDim ans(U_input_3, U_input_2, U_input_1)
For ii = 0 To U_input_1
    For j = 0 To U_input_2
        For k = 0 To U_input_3
            ans(k, j, ii) = input_3D_array(ii, j, k)
        Next k
    Next j
Next ii

rev_3D_array = ans
End Function

Function delete_firstIndex_3D_array(input3D_array, row_Index)
'функция , удаляющая определенный ряд (первый индекс) в трехмерном массиве
U_input_1 = UBound(input3D_array, 1)
U_input_2 = UBound(input3D_array, 2)
U_input_3 = UBound(input3D_array, 3)

ReDim ans(U_input_1 - 1, U_input_2, U_input_3)
first_ind = 0
For ii = 0 To U_input_1
    If ii <> row_Index Then
        For j = 0 To U_input_2
            For k = 0 To U_input_3
                ans(first_ind, j, k) = input3D_array(ii, j, k)
            Next k
        Next j
        first_ind = first_ind + 1
    End If
Next ii
delete_firstIndex_3D_array = ans
End Function



'ВЕКТОРЫ
Function VecMod(Vector)
'функция, вычисляющая модуль вектора
    calc = Sqr(Vector(0) ^ 2 + Vector(1) ^ 2 + Vector(2) ^ 2)
    VecMod = calc
End Function

Function VecProduct(Vector1, Vector2)
'функция, вычисляющая векторное произведение
Dim ans(2)
ans(0) = Vector1(1) * Vector2(2) - Vector1(2) * Vector2(1)
ans(1) = Vector1(2) * Vector2(0) - Vector1(0) * Vector2(2)
ans(2) = Vector1(0) * Vector2(1) - Vector1(1) * Vector2(0)
VecProduct = ans
End Function

Function two_points_NOT_overlap(point1, point2) As Boolean
'функция, которая подтверждает, что точки не совпадают
two_points_NOT_overlap = True

Xp1 = point1(0): Yp1 = point1(1): Zp1 = point1(2)
Xp2 = point2(0): Yp2 = point2(1): Zp2 = point1(2)

If Xp1 = Xp2 And Yp1 = Yp2 And Zp1 = Zp2 Then two_points_NOT_overlap = False

End Function

Function codirectional_of_vectors(some_point, other_point) As Boolean
'функция, проверяющая, что векторы
'1 (some_point-(0,0,0));
'2 (other_point-(0,0,0));
'все направлены в одну сторону
codirectional_of_vectors = False
cosV_on_Modules = vec_scalar_product(some_point, other_point)
If cosV_on_Modules > 0 Then codirectional_of_vectors = True

End Function

Function vec_scalar_product(vec1, vec2)
'функция, вычисляющая скалярное произведение векторов
vec_scalar_product = vec1(0) * vec2(0) + vec1(1) * vec2(1) + vec1(2) * vec2(2)
End Function

'3.ОСТАЛЬНЫЕ


Function point_is_inOrigin(point, precision) As Boolean
'функция, проверяющая, пренадлежит ли точка началу координат
'precision - точность в абсолютных величинах
point_is_inOrigin = False
If Abs(point(0)) < precision _
    And Abs(point(1)) < precision _
    And Abs(point(2)) < precision Then
        point_is_inOrigin = True
End If
End Function

Function find_Crosspoint_of_Ray_and_Plane(checking_vector, plane_norm, plane_point)
'функция поиска точки пересечения вектора и плоскости, заданной векторами
'!!! проверка - чтобы точка пересечения лежала именно на луче, а не прямой, т.е.
Dim ans(3)
'проверим сначала коллинеарны ли вектора плоскости
    'плоскость
    AA = plane_norm(0): BB = plane_norm(1): CC = plane_norm(2)
    DD = -(AA * plane_point(0) + BB * plane_point(1) + CC * plane_point(2))
    'точка пересечения исследуемого вектора и плоскости
    a0 = AA * checking_vector(0) + BB * checking_vector(1) + CC * checking_vector(2)
    If a0 <> 0 Then
            t = -DD / a0
            ans(3) = 0 'присваиваем 3ей координате код "0", который означает, что точка пересечения существует
            ans(0) = checking_vector(0) * t: ans(1) = checking_vector(1) * t: ans(2) = checking_vector(2) * t
    Else
            ans(3) = -1 'присваиваем 3ей координате код "-1", который означает, что a0=0
            ans(0) = 0: ans(1) = 0: ans(2) = 0
    End If
find_Crosspoint_of_Ray_and_Plane = ans
End Function

Function point_inside_triangle(P1, P2, P3, point) As Boolean
'!!!!!!надо сюда заложить точность определения и вычислений
'функция, которая определяет, лежит ли точка point внутри треугольника с вершинами P1,P2,P3
'площадь треугольника
S_full = triangle_square(P1, P2, P3)
'разбиваем треугольник на 3 треугольника и ищем площади
'1.point-P1-P2
S_point_P1_P2 = triangle_square(P1, P2, point)
'2.point-P1-P3
S_point_P1_P3 = triangle_square(P1, P3, point)
'3.point-P3-P2
S_point_P3_P2 = triangle_square(P3, P2, point)

If Abs(S_full - (S_point_P1_P2 + S_point_P1_P3 + S_point_P3_P2)) < 0.1 Then 'в формулу заложен критерий совпадения площадей в некотором интервале - от 0 до 0.1
    point_inside_triangle = True
Else
    point_inside_triangle = False
End If

End Function

Function triangle_square(P1, P2, P3)
'функция, которая ищет площадь треугольника по 3м вершинам, формула Герона
P1P2 = VecMod(Array(P1(0) - P2(0), P1(1) - P2(1), P1(2) - P2(2)))
P1P3 = VecMod(Array(P1(0) - P3(0), P1(1) - P3(1), P1(2) - P3(2)))
P2P3 = VecMod(Array(P2(0) - P3(0), P2(1) - P3(1), P2(2) - P3(2)))


'вычисляем полупериметр
P = (P1P2 + P1P3 + P2P3) / 2
If P * (P - P1P2) * (P - P1P3) * (P - P2P3) < 0 Then
    dfdf = 0
End If
'находим площадь по формуле Герона
triangle_square = Sqr(Abs(P * (P - P1P2) * (P - P1P3) * (P - P2P3)))

End Function

Function rad(dgree)
'функция преобразования угла градусов в радианы
Pi = Application.WorksheetFunction.Pi
rad = dgree * Pi / 180
End Function

Function deg(rdian)
'функция преобразования угла радиан в градусы
Pi = Application.WorksheetFunction.Pi
deg = rdian * 180 / Pi
End Function







