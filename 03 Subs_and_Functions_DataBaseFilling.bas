Attribute VB_Name = "Subs_and_Functions_IEKBD"

'������� � ��������� ��� ���������� ������ � ���������� ���� ������

'----------------------------------�������� � ����������� �����---------------------------------------------
'������
Sub Button_CreateLumCards() '������ ����������� ��������� ������ �� ������� �����������
'�������� ����� ��� �������������� ������
    srch_str = "Project_folder"
    Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
    start_RL = SearchCell.Row: start_colL = SearchCell.Column
    gen_path = ActiveSheet.Cells(start_RL, start_colL + 1)
    Dim path_folder As String
    For i = 1 To 7
        path_folder = gen_path & "\" & ActiveSheet.Cells(start_RL + i, start_colL + 1)
        If FolderExists(path_folder) = False Then MkDir path_folder
    Next i
'�������� ���������-�������� ������������
    srch_str = "LumNickname"
    Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
    start_RL = SearchCell.Row: start_colL = SearchCell.Column
    Lum_count = 0:
    Do While Worksheets("����� ������").Cells(start_RL + Lum_count + 1, start_colL).Text <> "" And Lum_count < 1000
        Worksheets("������").Copy before:=Worksheets("������")
        Lum_count = Lum_count + 1
        sheet_name = Worksheets("����� ������").Cells(start_RL + Lum_count, start_colL).Text
        Worksheets(Lum_count + 1).Name = sheet_name
        Fill_new_worksheet "����� ������", sheet_name, start_RL + Lum_count
    Loop
End Sub


Sub Button_FillDatabase() '������ ���������� ������������ ����� � ����� ��

'���������� ������������ ����� STRDBIK
srch_str = "Project_folder": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): start_RL = SearchCell.Row: start_colL = SearchCell.Column
Work_Path = ActiveSheet.Cells(start_RL, start_colL + 1)
work_ies_fold = ActiveSheet.Cells(start_RL + 1, start_colL + 1)
work_kss_image_fold = ActiveSheet.Cells(start_RL + 2, start_colL + 1)
work_lum_image_fold = ActiveSheet.Cells(start_RL + 3, start_colL + 1)
work_techdata_fold = ActiveSheet.Cells(start_RL + 4, start_colL + 1)
work_drawing_fold = ActiveSheet.Cells(start_RL + 5, start_colL + 1)
work_passport_fold = ActiveSheet.Cells(start_RL + 6, start_colL + 1)
work_CU_fold = ActiveSheet.Cells(start_RL + 7, start_colL + 1)

srch_str = "Control_file": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): start_RL = SearchCell.Row: start_colL = SearchCell.Column
strdbik_Path = Work_Path & "\" & ActiveSheet.Cells(start_RL, start_colL + 1)

sourceWbname = ThisWorkbook.Name '��� �������� ����� � �������� ����� ������������
Set sh_source = Workbooks(sourceWbname).Sheets(1)
'������� ���-�� ������� �� ������������
srch_str = "LumNickname": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): start_RL = SearchCell.Row: start_colL = SearchCell.Column
Lum_count = 0:
Do While Worksheets("����� ������").Cells(start_RL + Lum_count + 1, start_colL).Text <> "" And Lum_count < 1000
        Lum_count = Lum_count + 1
Loop
'�������� ������ ��� ���������� ����� �� � ����� ��
srch_str = "DataBase_Folder": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): full_db_path = ActiveSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
db_fold_name = Right(full_db_path, Len(full_db_path) - InStrRev(full_db_path, "\"))
ies_fold = ActiveSheet.Cells(SearchCell.Row + 1, SearchCell.Column + 1)
kss_image_fold = ActiveSheet.Cells(SearchCell.Row + 2, SearchCell.Column + 1)
lum_image_fold = ActiveSheet.Cells(SearchCell.Row + 3, SearchCell.Column + 1)
techdata_fold = ActiveSheet.Cells(SearchCell.Row + 4, SearchCell.Column + 1)
drawing_fold = ActiveSheet.Cells(SearchCell.Row + 5, SearchCell.Column + 1)
passport_fold = ActiveSheet.Cells(SearchCell.Row + 6, SearchCell.Column + 1)
CU_fold = ActiveSheet.Cells(SearchCell.Row + 7, SearchCell.Column + 1)



Workbooks.Open strdbik_Path '��������� ����������� ���� strdbik
Set sheet_strdbik = ActiveWorkbook.ActiveSheet

For i = 1 To Lum_count '���������
    If sh_source.Cells(start_RL + i, 47) = "" Then
        CU_f = ""
    Else
        CU_f = CU_fold
    End If
    'strdbik
    Fill_strdbik sheet_strdbik, sourceWbname, start_RL + i, sh_source.Cells(start_RL + i, 14), db_fold_name, ies_fold, _
    kss_image_fold, lum_image_fold, techdata_fold, drawing_fold, passport_fold, CU_f
    '������� �������� � ����������� �����
    lum_nickname = sh_source.Cells(start_RL + i, 5)
    sh_source.Cells(start_RL + i, 41) = lum_nickname & ".ies" '����� ies
    sh_source.Cells(start_RL + i, 42) = "c_" & lum_nickname & ".png" '��������� ���
    sh_source.Cells(start_RL + i, 43) = "i1_" & lum_nickname & ".png" '����������� �����������
    sh_source.Cells(start_RL + i, 44) = lum_nickname & ".txt" '�������� �����������
    sh_source.Cells(start_RL + i, 45) = "i2_" & lum_nickname & ".png" '������ �����������
    
Next i

'���������� ���� ������


Workbooks(sourceWbname).Activate 'sh_source
    
Copy_fromFolder_toFolder Work_Path & "\" & work_ies_fold & "\", full_db_path & "\" & ies_fold & "\" '����� IES
Copy_fromFolder_toFolder Work_Path & "\" & work_kss_image_fold & "\", full_db_path & "\" & kss_image_fold & "\" '��������� ���
Copy_fromFolder_toFolder Work_Path & "\" & work_lum_image_fold & "\", full_db_path & "\" & lum_image_fold & "\" '����������� �����������
Copy_fromFolder_toFolder Work_Path & "\" & work_techdata_fold & "\", full_db_path & "\" & techdata_fold & "\" '�������� �����������
Copy_fromFolder_toFolder Work_Path & "\" & work_drawing_fold & "\", full_db_path & "\" & drawing_fold & "\" '������ �����������
Copy_fromFolder_toFolder Work_Path & "\" & work_passport_fold & "\", full_db_path & "\" & passport_fold & "\" '�������� ������������
Copy_fromFolder_toFolder Work_Path & "\" & work_CU_fold & "\", full_db_path & "\" & CU_fold & "\" '�� �����������

End Sub

Sub Button_CheckFileExist_inDatabase() '������ ��� �������� - ��� �� ����� ����������� � ��
'����� ��
srch_str = "DataBase_Folder": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): full_db_path = ActiveSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
'db_fold_name = Right(full_db_path, Len(full_db_path) - InStrRev(full_db_path, "\"))
Dim gen_path(6)
ies_fold = ActiveSheet.Cells(SearchCell.Row + 1, SearchCell.Column + 1): gen_path(0) = full_db_path & "\" & ies_fold
kss_image_fold = ActiveSheet.Cells(SearchCell.Row + 2, SearchCell.Column + 1): gen_path(1) = full_db_path & "\" & kss_image_fold
lum_image_fold = ActiveSheet.Cells(SearchCell.Row + 3, SearchCell.Column + 1): gen_path(2) = full_db_path & "\" & lum_image_fold
techdata_fold = ActiveSheet.Cells(SearchCell.Row + 4, SearchCell.Column + 1): gen_path(3) = full_db_path & "\" & techdata_fold
drawing_fold = ActiveSheet.Cells(SearchCell.Row + 5, SearchCell.Column + 1): gen_path(4) = full_db_path & "\" & drawing_fold
passport_fold = ActiveSheet.Cells(SearchCell.Row + 6, SearchCell.Column + 1): gen_path(5) = full_db_path & "\" & passport_fold
CU_fold = ActiveSheet.Cells(SearchCell.Row + 7, SearchCell.Column + 1): gen_path(6) = full_db_path & "\" & CU_fold



'���� ������ �������:
srch_str = "LumNickname": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): i_start = SearchCell.Row


i = i_start + 1
Do While ActiveSheet.Cells(i, 41).Text <> "" And i < 1000
    For j = 41 To 47
        f_name = ActiveSheet.Cells(i, j)
        If f_name <> "" Then
            f_path = gen_path(j - 41) & "\" & f_name
            If Check_File_Exist(f_path) Then
                ActiveSheet.Cells(i, j).Interior.Color = xlNone
                ActiveSheet.Cells(i, j).Interior.ColorIndex = -4142
            Else
                ActiveSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0)
            End If
        End If
    Next j
i = i + 1
Loop


End Sub
Function Check_File_Exist(file_path) As Boolean
On Error Resume Next
Check_File_Exist = Dir(file_path) <> vbNullString
If Err.Number <> 0 Then Check_File_Exist = False
On Error GoTo 0
End Function

Sub Copy_fromFolder_toFolder(init_folder, distin_folder)
    Dim Shablon$, OnlyName$
    Shablon = "*.*"
    OnlyName = Dir(init_folder & Shablon, vbReadOnly + vbHidden + vbSystem)
    Do Until OnlyName = ""
        FileCopy init_folder & OnlyName, distin_folder & OnlyName
        OnlyName = Dir
    Loop
End Sub


Sub Fill_strdbik(Sh_strdbik, source_Wb_name, source_wbRow, short_series_name, _
                 db_folder, ies_folder, kss_image_folder, lum_image_folder, _
                 techdata_folder, drawing_folder, passport_folder, CU_folder) '��������� ���������� ������������ �����

    
'1. ���������� ������ ��� ������� (���� ����� ����� ��� ����, �� � ����� ���������, ���� ��� ����� �����, �� � ����� ����� � ������������ ����� �����)
    serie_exists = False
    goinOn = True
    i = 1
    Do While i < 1000 And goinOn
        '���� � ��������� ������ ����� (������ �����) �������, �� �� ����� �� ����� �����, ������, ���� ��������� ���� �������
        If serie_exists Then
            If Asc(Left(Sh_strdbik.Cells(i, 1), 1)) >= 192 And Asc(Left(Sh_strdbik.Cells(i, 1), 1)) <= 255 Or Sh_strdbik.Cells(i, 1).Text = "" Then
                Sh_strdbik.Cells(i, 1).EntireRow.Insert
                ins_Row_num = i
                goinOn = False
            End If
        End If
        '���� � ��������� ������ ��� ������ �� �� ����� �� ����� �������
        If serie_exists = False And Sh_strdbik.Cells(i, 1).Text = "" Then
            With Sh_strdbik.Cells(i, 1)
                .Value = short_series_name '���������� ����� �����
                .Font.Size = 12
                .Font.Bold = True
            End With
            Range(Sh_strdbik.Cells(i, 1), Sh_strdbik.Cells(i, 42)).Interior.Color = RGB(0, 255, 0) '����������� � ������� ����
            ins_Row_num = i + 1
            goinOn = False
        End If
        If Sh_strdbik.Cells(i, 1) = short_series_name Then serie_exists = True
        i = i + 1
    Loop
'2. ��������� ������� �������
Set Sh_sourcebook = Workbooks(source_Wb_name).Sheets(1)
lum_nickname = Sh_sourcebook.Cells(source_wbRow, 5)
With Sh_strdbik
    '----�������----
    For j = 2 To 26
        .Cells(ins_Row_num, j) = Sh_sourcebook.Cells(source_wbRow, 13 + j)
    Next j
    .Cells(ins_Row_num, 27) = Sh_sourcebook.Cells(source_wbRow, 6) '�������: ���-�� ��
    .Cells(ins_Row_num, 28) = Sh_sourcebook.Cells(source_wbRow, 8) '�������: ��������
    '----�������----
    
    .Cells(ins_Row_num, 1) = lum_nickname '������� �����������
    .Cells(ins_Row_num, 29) = Sh_sourcebook.Cells(source_wbRow, 3) '������� ������������
    .Cells(ins_Row_num, 30) = Sh_sourcebook.Cells(source_wbRow, 4) '�������
    .Cells(ins_Row_num, 37) = "���������� " & Sh_sourcebook.Cells(source_wbRow, 3) '������� ������������ ��� �������� IES
    .Cells(ins_Row_num, 39) = Sh_sourcebook.Cells(source_wbRow, 7) '�������� �����
    .Cells(ins_Row_num, 41) = Sh_sourcebook.Cells(source_wbRow, 10) '���������
    .Cells(ins_Row_num, 42) = Sh_sourcebook.Cells(source_wbRow, 9) '���� ������
    
    '----������-----
    .Cells(ins_Row_num, 31) = "\" & db_folder & "\" & passport_folder & "\" & Sh_sourcebook.Cells(source_wbRow, 46) '�������
    .Cells(ins_Row_num, 32) = "\" & db_folder & "\" & ies_folder & "\" & lum_nickname & ".ies" '���� IES
    .Cells(ins_Row_num, 33) = "\" & db_folder & "\" & kss_image_folder & "\" & "c_" & lum_nickname & ".png" '�������� ���
    .Cells(ins_Row_num, 34) = "\" & db_folder & "\" & lum_image_folder & "\" & "i1_" & lum_nickname & ".png" '����������� �����������
    .Cells(ins_Row_num, 35) = "\" & db_folder & "\" & drawing_folder & "\" & "i2_" & lum_nickname & ".png" '������ �����������
    .Cells(ins_Row_num, 36) = "\" & db_folder & "\" & techdata_folder & "\" & lum_nickname & ".txt"  '�������� �����������
    If CU_folder <> "" Then .Cells(ins_Row_num, 38) = "\" & db_folder & "\" & CU_folder & "\" & Sh_sourcebook.Cells(source_wbRow, 47) '������� ������������� �������������
    .Cells(ins_Row_num, 40) = "\" & db_folder & "\" & passport_folder & "\" & Sh_sourcebook.Cells(source_wbRow, 46)  '���� ��� ��� ������ - ��������, � �� ����� ��� ������ - ����� ���� ����� �������
    
End With

End Sub



Sub Fill_new_worksheet(base_w_name, new_w_name, basetable_Row) '��������� ���������� ����� ���������� �����
Set BaseWS = Worksheets(base_w_name)
Set NewWS = Worksheets(new_w_name)

'0. ���������
NewWS.Cells(1, 2) = BaseWS.Cells(basetable_Row, 3)

'1.������ � ��������� IES-�����
srch_str = "start IES file": Set SearchCell = NewWS.Cells.Find(what:=srch_str, searchformat:=False): S_Row = SearchCell.Row: S_Col = SearchCell.Column

With NewWS
    .Cells(S_Row + 4, S_Col + 2) = BaseWS.Cells(basetable_Row, 4) '�������
    .Cells(S_Row + 5, S_Col + 2) = BaseWS.Cells(basetable_Row, 2) '������������
    .Cells(S_Row + 10, S_Col + 1) = BaseWS.Cells(basetable_Row, 6) '���-�� ��
    .Cells(S_Row + 10, S_Col + 2) = BaseWS.Cells(basetable_Row, 7) '�v
    .Cells(S_Row + 10, S_Col + 8) = BaseWS.Cells(basetable_Row, 11) / 1000 '������� � - ��������� � [�]
    .Cells(S_Row + 10, S_Col + 9) = BaseWS.Cells(basetable_Row, 12) / 1000 '������� � - ��������� � [�]
    .Cells(S_Row + 10, S_Col + 10) = BaseWS.Cells(basetable_Row, 13) / 1000 '������� � - ��������� � [�]
    .Cells(S_Row + 11, S_Col + 3) = BaseWS.Cells(basetable_Row, 8) '��������
End With

'2.������ � ��������� ����� ����������� ����������
srch_str = "start_tech_data": Set SearchCell = NewWS.Cells.Find(what:=srch_str, searchformat:=False): S_Row = SearchCell.Row: S_Col = SearchCell.Column

With BaseWS
'�������� ������� ���������� ��������
    gab_razm = Abs(.Cells(basetable_Row, 11)) & "x" & Abs(.Cells(basetable_Row, 12)) & "x" & .Cells(basetable_Row, 13)
End With
With NewWS
    .Cells(S_Row + 2, S_Col + 1) = BaseWS.Cells(basetable_Row, 3) '������� ������������
    .Cells(S_Row + 3, S_Col + 1) = BaseWS.Cells(basetable_Row, 4) '�������
    .Cells(S_Row + 5, S_Col + 1) = BaseWS.Cells(basetable_Row, 8) '��������
    .Cells(S_Row + 6, S_Col + 1) = gab_razm '���������� �������
    .Cells(S_Row + 7, S_Col + 1) = BaseWS.Cells(basetable_Row, 7) '�v
End With

End Sub

Function FolderExists(ByRef path As String) As Boolean '�������, ������� ��������� ������������� �����
On Error Resume Next
FolderExists = GetAttr(path)
End Function



'----------------------------------�������� � ������� ��������� ����� �����������-----------------------------------------------------------------------
'������
Sub Button_Data_frmPrepare_toCreateArea() '���������� ������, ���������� �������������� �������� ��� (1 ����)
    Data_frmPrepare_toCreateArea
End Sub

Sub Button_Refresh_2DPolar() '���������� �������� �������� ��� (�� 1��� �����)
    '����� ������� ��������, �� ������� ����� ����������
    srch_str = "edit KSS"
    Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
    start_RE = SearchCell.Row: start_colE = SearchCell.Column
    edit_Gamma_count = 0
    Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
        edit_Gamma_count = edit_Gamma_count + 1
    Loop
    
    
    
    '���������� �������� �������� ���
    Dim gamma_Range As Range, I1_Range As Range, I2_Range As Range
    Set gamma_Range = ActiveSheet.Range(Cells(start_RE, start_colE + 1), Cells(start_RE, start_colE + edit_Gamma_count))
    Set I1_Range = ActiveSheet.Range(Cells(start_RE + 1, start_colE + 1), Cells(start_RE + 1, start_colE + edit_Gamma_count))
    Set I2_Range = ActiveSheet.Range(Cells(start_RE + 2, start_colE + 1), Cells(start_RE + 2, start_colE + edit_Gamma_count))
    
    built_linear_KSS "ChartRect0", gamma_Range, I1_Range
    built_linear_KSS "ChartRect90", gamma_Range, I2_Range
    
    '���������� �������� �������� ���
    build_2polar_KSS "ChartPolar1", gamma_Range, I1_Range, I2_Range, 350, 350, 48, 6, 1
    
End Sub
Sub Button_Refresh_IESOutput() '��������� ���������� ������� IES
'�������� ������ ��� ����������

'������ ���������� ����� (gamma � �)  � ������� ������������� ���, �� ������ ������� ����� ����������� IES ����
srch_str = "edit KSS"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)

edit_Gamma_count = 0: 'edit_C_count = 0

start_RE = SearchCell.Row: start_colE = SearchCell.Column
Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
    edit_Gamma_count = edit_Gamma_count + 1
Loop


ReDim G(edit_Gamma_count - 1), Ic0(edit_Gamma_count - 1), Ic90(edit_Gamma_count - 1)
For j = 1 To edit_Gamma_count
    G(j - 1) = Cells(start_RE, start_colE + j)
    Ic0(j - 1) = Cells(start_RE + 1, start_colE + j)
    Ic90(j - 1) = Cells(start_RE + 2, start_colE + j)  ' ���� ����� 2 �������
Next j
    
    
'������� ������� ������� !!! ���� �� ������� �������� ���������� �����, ���� ���-�� ����� "�" ������ 19
srch_str = "start IES matrix"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
start_RE = SearchCell.Row: start_colE = SearchCell.Column
'������� ��������������� ������ �������
edit_Gamma_count = 0: edit_C_count = 0

Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
    edit_Gamma_count = edit_Gamma_count + 1
Loop
Do While ActiveSheet.Cells(start_RE + edit_C_count + 1, start_colE + 1).Text <> "" And edit_C_count < 1000
    edit_C_count = edit_C_count + 1
Loop

'������� ���������� �������� � ����
If edit_Gamma_count > 0 And edit_C_count > 0 Then
    Range(Cells(start_RE - 1, start_colE + 1), Cells(start_RE + edit_C_count, start_colE + edit_Gamma_count)).Clear
    Range(Cells(start_RE + 1, start_colE), Cells(start_RE + edit_C_count, start_colE)).Clear '������� � ���� ������� �����
End If

N_c = ActiveSheet.Cells(start_RE - 3, start_colE + 5) '������ �� ������, � ������� ���-�� ����� "�"

Fill_iesFile_Output start_RE - 1, start_colE, G, N_c, Ic0, Ic90
    
    
End Sub

Sub Button_write_IES_file()
'���������� ����/��� ������������� �����
Set GenSheet = ThisWorkbook.Sheets(1)

srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = GenSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
srch_str = "����� ies": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = GenSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
ies_name = ActiveSheet.Name & ".ies"

ies_path = gen_path & "\" & Dir_path & "\" & ies_name

ind1 = Convert_ShapeName_to_ShapeNumber("RB_IESfile_New")
ind2 = Convert_ShapeName_to_ShapeNumber("RB_IESfile_Existing")
If ActiveSheet.Shapes(ind1).ControlFormat.Value > 0 Then 'mess = "�����"
    srch_str = "start IES file": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
    create_IES_file ies_path, SearchCell.Row, SearchCell.Column + 1, SearchCell.Row + 12
    MsgBox "OK!"
End If
If ActiveSheet.Shapes(ind2).ControlFormat.Value > 0 Then 'mess = "������������"
    existing_IES_path = GetFilePath(, gen_path)
    Copy_File existing_IES_path, ies_path
End If

'��������� � ���������� �������� ���
'����� ������� ��������, �� ������� ����� ����������
    srch_str = "edit KSS"
    Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
    start_RE = SearchCell.Row: start_colE = SearchCell.Column
    edit_Gamma_count = 0
    Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
        edit_Gamma_count = edit_Gamma_count + 1
    Loop
    
    
    
    '�������� ������
    Dim gamma_Range As Range, I1_Range As Range, I2_Range As Range
    Set gamma_Range = ActiveSheet.Range(Cells(start_RE, start_colE + 1), Cells(start_RE, start_colE + edit_Gamma_count))
    Set I1_Range = ActiveSheet.Range(Cells(start_RE + 1, start_colE + 1), Cells(start_RE + 1, start_colE + edit_Gamma_count))
    Set I2_Range = ActiveSheet.Range(Cells(start_RE + 2, start_colE + 1), Cells(start_RE + 2, start_colE + edit_Gamma_count))

    '����� �������, ��� ����� �������� ������ ���
    srch_str = "1. ���": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): start_RE = SearchCell.Row: start_colE = SearchCell.Column

'���������� ���
'Dim pic_H, pic_W As Double
pic_W = 340: pic_H = 340
build_2polar_KSS "OutputChartPolar", gamma_Range, I1_Range, I2_Range, 340, 340, start_RE + 3, start_colE + 0, 1


'������� ���
srch_str = "��������� ���": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = GenSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
AS_name = ActiveSheet.Name
pic_name = "c_" & ActiveSheet.Name & ".png"
export_pic_path = gen_path & "\" & Dir_path & "\" & pic_name
    
If Val(Application.Version) > 11 Then
    Dim chtChart As Chart
    Set chtChart = Charts.Add
    For Each s In chtChart.SeriesCollection
        s.Delete
    Next s
    Set chtChart = chtChart.Location(Where:=xlLocationAsObject, Name:=AS_name)
    chtChart.ChartArea.Border.Color = RGB(255, 255, 255)
    chtChart.Parent.Height = pic_H
    chtChart.Parent.Width = pic_W
    
    num = Convert_ShapeName_to_ShapeNumber("OutputChartPolar")
    ActiveSheet.Shapes(num).Copy
    'chtChart.Activate
    With ActiveChart
            .ChartArea.Select
            .Paste
    End With
    ActiveChart.Export Filename:=export_pic_path, filtername:="PNG" '��������� ��������
    ActiveChart.Parent.Delete
Else
    ActiveChart.Export Filename:=export_pic_path, filtername:="PNG" '��������� ��������
End If

MsgBox "�������� ���������"
End Sub

Sub Button_load_and_save_LumImage()
'������� ������
    'pic_path = "d:\������\��� 2016\21 03 2016 ������������� �������� IES\!!�� ������ ������\c_sdo01_20.png"
    '����
    srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "�������� ������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): InitialData_Path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "�����������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    pic_name = "i1_" & ActiveSheet.Name & ".png"
    export_pic_path = gen_path & "\" & Dir_path & "\" & pic_name
    pic_path = GetFilePath(, InitialData_Path)
    diagramm_N = Convert_ChartName_to_ChartNumber("OutputLumPhoto")
    'export_path = "D:/����������/1.gif": file_ext = "GIF"
    take_and_save_picture pic_path, diagramm_N, 400, 400, export_pic_path, "PNG"
End Sub
Sub Button_load_and_save_LumDrawing()
'������� ������
    'pic_path = "d:\������\��� 2016\21 03 2016 ������������� �������� IES\!!�� ������ ������\c_sdo01_20.png"
    '����
    srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "�������� ������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): InitialData_Path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "�������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    pic_name = "i2_" & ActiveSheet.Name & ".png"
    export_pic_path = gen_path & "\" & Dir_path & "\" & pic_name
    pic_path = GetFilePath(, InitialData_Path)
    diagramm_N = Convert_ChartName_to_ChartNumber("OutputLumDrawing")
    'export_path = "D:/����������/1.gif": file_ext = "GIF"
    take_and_save_picture pic_path, diagramm_N, 800, 800, export_pic_path, "PNG"
End Sub

Sub Button_save_Passport()
    srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "�������� ������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): InitialData_Path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "��������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "lum_passport": Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False): passport_name = ActiveSheet.Cells(SearchCell.Row, SearchCell.Column + 1)
    passport_name = passport_name & ".pdf"
    
    source_passport_path = GetFilePath(, InitialData_Path)
    save_passport_path = gen_path & "\" & Dir_path & "\" & passport_name
    
    '�������� ����
    Copy_File source_passport_path, save_passport_path
    '������ ������ � �������� �������
    lum_nickname = ActiveSheet.Name
    srch_str = lum_nickname: Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False)
    Sheets(1).Cells(SearchCell.Row, 46) = passport_name
End Sub
Sub Button_Write_TechData_File()
    srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "��� ������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    td_filename = ActiveSheet.Name & ".txt"

    td_path = gen_path & "\" & Dir_path & "\" & td_filename
    
    Create_TechData_File td_path
    
    MsgBox "Ok!"
End Sub

Sub Button_Write_CU_File()
    srch_str = "Project_folder": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): gen_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    srch_str = "����� �������������": Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False): Dir_path = Sheets(1).Cells(SearchCell.Row, SearchCell.Column + 1)
    cu_filename = "cu_" & ActiveSheet.Name & ".txt"

    cu_path = gen_path & "\" & Dir_path & "\" & cu_filename
    
    '������ ������ � �������� �������
    lum_nickname = ActiveSheet.Name
    srch_str = lum_nickname: Set SearchCell = Sheets(1).Cells.Find(what:=srch_str, searchformat:=False)
    Sheets(1).Cells(SearchCell.Row, 47) = cu_filename

    '��� ����������� � �������� �����
    L_short_name = Sheets(1).Cells(SearchCell.Row, 3)
    L_Flux = Sheets(1).Cells(SearchCell.Row, 7)
    
    Create_CoeffUse_File cu_path, 3, L_short_name, L_Flux
    
    MsgBox "Ok!"
End Sub

Sub FillRangeWithInterpolatedData()
expAnglesExcel = Application.InputBox("����������������� ������. �������� ����:", Type:=64)
expDataExcel = Application.InputBox("����������������� ������. �������� ��������:", Type:=64)
OutputAngles = Application.InputBox("�������� ������. �������� ����:", Type:=64)
Set OutputStartCell = Application.InputBox("�������� ������. �������� ������, ������� � ������� ��������� ��������� ������:", Type:=8) '

ReDim expAngles(UBound(expAnglesExcel) - 1)
ReDim expData(UBound(expDataExcel) - 1)
For i = 0 To UBound(expAnglesExcel) - 1
    expAngles(i) = expAnglesExcel(i + 1, 1)
    expData(i) = expDataExcel(i + 1, 1)
Next i

CountOutput = UBound(OutputAngles): LastExperimental = UBound(expAngles)
OutputStartRow = OutputStartCell.Row: OutputStartColumn = OutputStartCell.Column
'������ � ��������� ��������
ActiveSheet.Cells(OutputStartRow, OutputStartColumn) = expData(0)
ActiveSheet.Cells(OutputStartRow + CountOutput - 1, OutputStartColumn) = expData(LastExperimental)
For i = 2 To CountOutput - 1
    x0 = OutputAngles(i, 1)
    indices_x1x2 = findElementInArray(x0, expAngles): i1 = indices_x1x2(0): i2 = indices_x1x2(1)
    If i1 = i2 Then
        ActiveSheet.Cells(OutputStartRow + i - 1, OutputStartColumn) = expData(i1)
    Else
        x1 = expAngles(i1): x2 = expAngles(i2)
        y1 = expData(i1): y2 = expData(i2)
        ActiveSheet.Cells(OutputStartRow + i - 1, OutputStartColumn) = arythm_interp(x0, x1, x2, y1, y2)
    End If
Next i


End Sub




Function findElementInArray(fElement, fArray) As Variant() '�������, ������� ���������� ��������� �������� � �������
Dim ans(1)
ans(0) = 0: ans(1) = 1
For i = 0 To UBound(fArray) - 1
    If fArray(i) - fElement < 0 And fArray(i + 1) - fElement > 0 Then
        ans(0) = i: ans(1) = i + 1
    End If
    If fArray(i) - fElement = 0 Then
        ans(0) = i: ans(1) = i
    End If
    If fArray(i + 1) - fElement = 0 Then
        ans(0) = i + 1: ans(1) = i + 1
    End If
Next i
findElementInArray = ans
End Function

Sub Create_CoeffUse_File(tfname, fluct_percent, lum_name, lum_flux) '���������, ������� ���������� � ���� ������������ �������������, ������������ �� ������ ������� � ������������
'fluct_percent - ���������� � ���������
'���� ��������� ����� ��� ���������� ����������
srch_str = "start_CU_data"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
start_R = SearchCell.Row: start_col = SearchCell.Column

Dim CU_list(12)
CU_list(0) = lum_name & Chr(9) & lum_flux
For i = 1 To 12
cu_str = ActiveSheet.Cells(start_R + i, start_col)
    For j = 1 To 7
        If i = 1 Then
            cu_str = cu_str & vbTab & ActiveSheet.Cells(start_R + i, start_col + j) 'vbTab
        Else
            'If j = 0 Then
              '  cu_str = cu_str & Chr(9) & ActiveSheet.Cells(start_R + i, start_col + j)
            'Else
                cu_strnum = Round(ActiveSheet.Cells(start_R + i, start_col + j) * (1 + fluct_percent * Rnd / 100), 3)
                cu_str = cu_str & Chr(9) & Str(cu_strnum)
            'End If
        End If
            
    Next j
    CU_list(i) = cu_str
Next i

FullPath = tfname

Open FullPath For Output As #1

'���������� � ����
For i = 0 To UBound(CU_list) - 1
    Print #1, CU_list(i) 'string_output(i)
Next i
Print #1, CU_list(UBound(CU_list)); ' ";" � ����� �������� ������� �������

Close #1

End Sub

Sub Create_TechData_File(tfname)

'���� ��������� ����� ��� ���������� ����������
srch_str = "start_tech_data"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
start_R = SearchCell.Row: start_col = SearchCell.Column
'�������� ���������� �� �����������
ReDim Tech_Info(6)
Tech_Info(0) = ActiveSheet.Cells(start_R + 1, start_col + 1) '������������ �����
For i = 1 To 6
    Tech_Info(i) = ActiveSheet.Cells(start_R + 1 + i, start_col) & " " & ActiveSheet.Cells(start_R + 1 + i, start_col + 1)
Next i
'��� ���� �� ���������� ����
s_number = Convert_ShapeName_to_ShapeNumber("text_extra_info")
Set textField = ActiveSheet.Shapes(s_number).DrawingObject
Extra_Info = Split(textField.Text, vbLf) '���� �� ��������� ��������� ������� - ����� ����������� vbCr ��� vbCrLf

'��������� ��������� ������ �� �����
ReDim Preserve Tech_Info(6 + UBound(Extra_Info) + 1)
For i = 7 To UBound(Tech_Info)
    Tech_Info(i) = Extra_Info(i - 7)
Next i


FullPath = tfname
Open FullPath For Output As #1

'���������� � ����
For i = 0 To UBound(Tech_Info)
    Print #1, Tech_Info(i)
Next i

Close #1
End Sub

'1 ����. ���������� ������, ���������� �������������� �������� ���

Sub Data_frmPrepare_toCreateArea()
'1) ������ ��������� ��������� �������� ��� �� ������� ���������� �������� � ������� ������������ IES-�����
'2) ��������� ������� ������� ���������� ����� ��������� ����, ������� �������� ���������� �� ������� ���������; ���������� ��������������� �����������
'3) ��������� ������� ������������ IES-����� ���� ��������� ����, ������� �������� ���������� �� ������� ���������; ���������� ��������������� �����������,
'�������� ��� � ������� ���������� �����������, ���� ��� �������� ���� ����� ���� ���������� ������
'4) ��������� �� ���������� ������� ������� ������������ IES-�����, ���� � ������� ���������� ��� ���-�� ����� "�" �����������


'������ ���������� ����� (gamma � �)  � ������� ������������� ���, �� ������ ������� ����� ����������� IES ����
srch_str = "edit KSS"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)

edit_Gamma_count = 0: edit_C_count = 0

start_RE = SearchCell.Row: start_colE = SearchCell.Column
Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
    edit_Gamma_count = edit_Gamma_count + 1
Loop
Do While ActiveSheet.Cells(start_RE + edit_C_count + 1, start_colE + 1).Text <> "" And edit_C_count < 1000
    edit_C_count = edit_C_count + 1
Loop

'������� ���������� �������� � ����
Range(Cells(start_RE, start_colE + 1), Cells(start_RE + edit_C_count, start_colE + edit_Gamma_count)).Clear
Range(Cells(start_RE + 1, start_colE), Cells(start_RE + edit_C_count, start_colE)).Clear '������� � ���� ������� �����



'������ ���������� ����� (gamma � C) � ������� ����������������� (��������������� ������)
srch_str = "prepare start"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)

Gamma_count = 0: C_count = 0

start_R = SearchCell.Row: start_col = SearchCell.Column
'���� �����
Do While ActiveSheet.Cells(start_R + Gamma_count + 1, start_col + 1).Text <> "" And Gamma_count < 1000
    Gamma_count = Gamma_count + 1
Loop
'���� �
Do While ActiveSheet.Cells(start_R, start_col + C_count + 2).Text <> "" And C_count < 1000
    C_count = C_count + 1
Loop


'������� ������� � ���� ��������, ����� ������ ���������� ������, �� �������� �� ����� �������
d_Gamma = Gamma_count - edit_Gamma_count: d_Row = C_count - 2

'�������� �����
'���� ������� d_Gamma ������ ����, �� ������� �� d_Gamma ������. ���� ������� �� ���� ������ 4� (� ������, ���� � ������� "edit" ���-�� ����� ����� <19), �� �������� �� ����������;
'���� ������� d_Gamma ������ ����, �� ������� �� ��� ����� - �� �� ����� 25��� �������

If d_Gamma > 0 Then

If edit_Gamma_count < 19 Then d_Gamma = d_Gamma - (19 - edit_Gamma_count)
    
    For i = 1 To d_Gamma
        Cells(1, start_col).EntireColumn.Insert
    Next i

End If

If d_Gamma < 0 Then
    If start_col - Abs(d_Gamma) < 25 Then d_Gamma = start_col - 25
    For i = 1 To Abs(d_Gamma)
        Cells(1, start_col - i).EntireColumn.Delete
    Next i
End If



'��������� �����������  � ����������
srch_str = "1. ����������"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
ActiveSheet.Hyperlinks.Add anchor:=SearchCell, Address:="", SubAddress:=Cells(1, d_Gamma + 54).Address, TextToDisplay:=srch_str

'�������� ���� � ������� ����������� ��������� ��� ������������ ������� ��� ���� ��� �������, ��� ������ ����� � �� ����������� ��������� �������� ������ 2� ��������
srch_str = "edit KSS"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
SC_Row = SearchCell.Row
For i = 1 To d_Row
    Cells(SC_Row + 1, 1).EntireRow.Insert
Next i

'��������� ����������� � ����������
srch_str = "2. ������������ IES-�����"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
ActiveSheet.Hyperlinks.Add anchor:=SearchCell, Address:="", SubAddress:=Cells(d_Row + 126, 1).Address, TextToDisplay:=srch_str
srch_str = "3. ����������� � ��������"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
ActiveSheet.Hyperlinks.Add anchor:=SearchCell, Address:="", SubAddress:=Cells(d_Row + 176, 1).Address, TextToDisplay:=srch_str

'� ������� "����������" ���������� ������� �������� ��� �� ����� �����, ���� � ���������� �������� ��� "������ ����"
'- ��� ��������, ���� ���-�� ��� �� ����� ������ 75

If Gamma_count >= 75 And d_Row > 0 Then
    Range(Cells(SC_Row + d_Row + 1, start_col + d_Gamma + 1), Cells(start_R + dRow + Gamma_count + 1, start_col + C_count + d_Gamma + 1)).Select
    Selection.Cut
    Cells(SC_Row + 1, start_col + d_Gamma + 1).Activate
    ActiveSheet.Paste
End If



'���������� ������� ������ "�����" � ���������� ���
ReDim G(Gamma_count - 1), Ic0(Gamma_count - 1), Ic90(Gamma_count - 1)
For i = 1 To Gamma_count
    G(i - 1) = Cells(start_R + i, start_col + d_Gamma + 1 + 0)
    Ic0(i - 1) = Cells(start_R + i, start_col + d_Gamma + 1 + 1)
    Ic90(i - 1) = Cells(start_R + i, start_col + d_Gamma + 1 + C_count) '������ ��������� ������� ��� - ���� ���
    For j = 0 To C_count
        Cells(start_RE + j, start_colE + i) = Cells(start_R + i, start_col + d_Gamma + 1 + j)
    Next j
Next i
'�������: ���� �
For j = 1 To C_count
    Cells(start_RE + j, start_colE) = Cells(start_R, start_col + d_Gamma + 1 + j)
Next j

'���������� �������� �������� ���
Dim gamma_Range As Range, I1_Range As Range, I2_Range As Range
Set gamma_Range = ActiveSheet.Range(Cells(start_RE, start_colE + 1), Cells(start_RE, start_colE + Gamma_count))
Set I1_Range = ActiveSheet.Range(Cells(start_RE + 1, start_colE + 1), Cells(start_RE + 1, start_colE + Gamma_count))
Set I2_Range = ActiveSheet.Range(Cells(start_RE + 2, start_colE + 1), Cells(start_RE + 2, start_colE + Gamma_count))

built_linear_KSS "ChartRect0", gamma_Range, I1_Range
built_linear_KSS "ChartRect90", gamma_Range, I2_Range

'���������� �������� �������� ���
build_2polar_KSS "ChartPolar1", gamma_Range, I1_Range, I2_Range, 350, 350, 48, 6, 1

'������ IES
srch_str = "start IES matrix"
Set SearchCell = ActiveSheet.Cells.Find(what:=srch_str, searchformat:=False)
start_RE = SearchCell.Row: start_colE = SearchCell.Column
'������� ��������������� ������ �������
edit_Gamma_count = 0: edit_C_count = 0

Do While ActiveSheet.Cells(start_RE, start_colE + edit_Gamma_count + 1).Text <> "" And edit_Gamma_count < 1000
    edit_Gamma_count = edit_Gamma_count + 1
Loop
Do While ActiveSheet.Cells(start_RE + edit_C_count + 1, start_colE + 1).Text <> "" And edit_C_count < 1000
    edit_C_count = edit_C_count + 1
Loop

'������� ���������� �������� � ����
If edit_Gamma_count > 0 And edit_C_count > 0 Then
    Range(ActiveSheet.Cells(start_RE, start_colE + 1), ActiveSheet.Cells(start_RE + edit_C_count, start_colE + edit_Gamma_count)).Clear
    Range(ActiveSheet.Cells(start_RE + 1, start_colE), ActiveSheet.Cells(start_RE + edit_C_count, start_colE)).Clear '������� � ���� ������� �����
End If

N_c = ActiveSheet.Cells(start_RE - 3, start_colE + 5) '������ �� ������, � ������� ���-�� ����� "�"
Fill_iesFile_Output start_RE - 1, start_colE, G, N_c, Ic0, Ic90


End Sub
Sub built_linear_KSS(ch_name As String, rX As Range, rY As Range)
'���������, �������� �������� ������ ���
Set AS1 = ActiveSheet
For s = 1 To AS1.ChartObjects.Count
    If AS1.ChartObjects(s).Name = ch_name Then
        
        AS1.ChartObjects(s).Activate
        Set chtChart = ActiveChart
        For Each s1 In chtChart.SeriesCollection
            s1.Delete
        Next s1
        '��������� �������
        chtChart.ChartType = xlXYScatter 'LinesNoMarkers
        
        chtChart.SeriesCollection.NewSeries
        With chtChart.SeriesCollection(1)
            .XValues = rX
            .Values = rY
        End With
    End If
Next s

End Sub
Sub build_2polar_KSS(ch_name As String, ang_R As Range, I_R1 As Range, I_R2 As Range, ch_H As Double, ch_W As Double, cell_R As Integer, cell_C As Integer, Frame_mode As Integer)
'��������� ���������� �������� ���� ��� � �������� �����������
'������ �������� - ��� �������, ���� � ����� ������ ������������ ���������, �� ��������� �� ��������, ���� �� ������������, �� ������ �����
'����� ���������� - ��� ���� � �������� �������; ������-������; ���-������� ��� �������
'Frame_mode - ����� ������������ ����� �������: 1 - ����� ���������� �������; 2-����� �������� ���, ����� ��� ����������� ��������� ������

Set AS1 = ActiveSheet

'01. �������� ���������� ��������� ��� ����������� � ����� �� ������
'�������� ����������� ����������� ��� ���������
s = 1
Do While s <= AS1.Shapes.Count
        If AS1.Shapes(s).Type = msoPicture And AS1.Shapes(s).Name = ch_name Then AS1.Shapes(s).Delete
        s = s + 1
Loop ' s

For s = 1 To AS1.ChartObjects.Count
    If AS1.ChartObjects(s).Name = ch_name Then AS1.ChartObjects(s).Delete
Next s

'02. ����������� �����, ��� ����� �������� ��� ����� ���������
N_r = (2 + 4 + 13) * 2: N_c = ang_R.Count '(2 + 4 + 7)*2
r1 = find_empty_area(cell_R, N_r, 2 * N_c - 1)


'03. ������ � ������ ��������
row_ang = ang_R.Row: col_ang = ang_R.Column '��� � �������, � ������� ���������� ����
row_I1 = I_R1.Row: col_I1 = I_R1.Column '��� � �������, � ������� ���������� ��� 1
row_I2 = I_R2.Row: col_I2 = I_R2.Column '��� � �������, � ������� ���������� ��� 2

'��������� �������� � ���������� �� � ��������� ������� �����
For j = 0 To 2 * N_c - 2
    If j < 18 Then
        ang = ActiveSheet.Cells(row_ang, col_ang + (N_c - 1) - j) * 3.14159 / 180
        Ic = ActiveSheet.Cells(row_I1, col_I1 + (N_c - 1) - j)
        I_X1 = -Ic * Sin(ang): I_Y1 = -Ic * Cos(ang)
        Ic = ActiveSheet.Cells(row_I2, col_I2 + (N_c - 1) - j)
        I_X2 = -Ic * Sin(ang): I_Y2 = -Ic * Cos(ang)
    Else
        ang = ActiveSheet.Cells(row_ang, col_ang + j - (N_c - 1)) * 3.14159 / 180
        Ic = ActiveSheet.Cells(row_I1, col_I1 + j - (N_c - 1))
        I_X1 = Ic * Sin(ang): I_Y1 = -Ic * Cos(ang)
        Ic = ActiveSheet.Cells(row_I2, col_I2 + j - (N_c - 1))
        I_X2 = Ic * Sin(ang): I_Y2 = -Ic * Cos(ang)
    End If
    
    ActiveSheet.Cells(r1, j + 1) = I_X1: ActiveSheet.Cells(r1 + 1, j + 1) = I_Y1
    ActiveSheet.Cells(r1 + 2, j + 1) = I_X2: ActiveSheet.Cells(r1 + 3, j + 1) = I_Y2
Next j

'04. ������� ����� ���������
'������ ��� ��� �����������
Set Xrange1 = ActiveSheet.Range(AS1.Cells(r1, 1), AS1.Cells(r1, 2 * N_c - 1))
Set datarange1 = ActiveSheet.Range(AS1.Cells(r1 + 1, 1), AS1.Cells(r1 + 1, 2 * N_c - 1))
Set Xrange2 = ActiveSheet.Range(AS1.Cells(r1 + 2, 1), AS1.Cells(r1 + 2, 2 * N_c - 1))
Set datarange2 = ActiveSheet.Range(AS1.Cells(r1 + 3, 1), AS1.Cells(r1 + 3, 2 * N_c - 1))

'������ �����
Imax = max_from_2(Application.Max(I_R1), Application.Max(I_R2))


'����������:
'���������� ���������� ��������� �������
Dim No(3)
For o = 0 To 3
    For i = 0 To 36
        ang = 3.14159 + 10 * 3.14159 / 180 * i
        Xo = (1 - o / 4) * Imax * Cos(ang)
        Yo = -(1 - o / 4) * Imax * Sin(ang)
    ActiveSheet.Cells(r1 + 4 + o * 2, i + 1) = Xo
    ActiveSheet.Cells(r1 + 4 + o * 2 + 1, i + 1) = Yo
    Next i
    
Next o

'�����:
For l = 0 To 12 '6
    ang = 3.14159 + 15 * 3.14159 / 180 * l
        Xo = (Imax) * Cos(ang)
        Yo = (Imax) * Sin(ang)
    ActiveSheet.Cells(r1 + 12 + l * 2, 1) = -Xo: ActiveSheet.Cells(r1 + 12 + l * 2, 2) = Xo
    ActiveSheet.Cells(r1 + 12 + l * 2 + 1, 1) = -Yo: ActiveSheet.Cells(r1 + 12 + l * 2 + 1, 2) = Yo
Next l

'----------�����-------------------------------

sheetname = ActiveSheet.Name






Dim chtChart As Chart
Set chtChart = Charts.Add
With chtChart.ChartArea.Border
    '.Weight = xlHairline
    '.LineStyle = xlLineStyleNone
    .Color = RGB(255, 255, 255)
End With

'��� 2003 excel ������ ����������� �� ��������� ������
'If Val(Application.Version) = 11 Then
    For Each s In chtChart.SeriesCollection
        s.Delete
    Next s
'End If
'��������� ������� ������� (I90)
chtChart.SeriesCollection.NewSeries
With chtChart.SeriesCollection(1)
    .XValues = Xrange2
    .Values = datarange2
    With .Border
        .Color = RGB(205, 175, 149) '�������
        .Weight = 4
    End With
End With
'��������� �������� ������� (I0)
chtChart.SeriesCollection.NewSeries
With chtChart.SeriesCollection(2)
    .XValues = Xrange1
    .Values = datarange1
    With .Border
        .Color = RGB(255, 0, 0) ' RGB(255, 0, 0) - �������
        .Weight = 4
    End With
End With

With chtChart.PlotArea.Border
    '.Weight = xlHairline
    '.LineStyle = 1
    .Color = RGB(255, 255, 255)
End With


'��������� �������� ���������� �����
For o = 0 To 3

    Set Xrange_O = Sheets(sheetname).Range(AS1.Cells(r1 + 4 + o * 2, 1), AS1.Cells(r1 + 4 + o * 2, 37))
    Set Yrange_O = Sheets(sheetname).Range(AS1.Cells(r1 + 4 + o * 2 + 1, 1), AS1.Cells(r1 + 4 + o * 2 + 1, 37))
    
    chtChart.SeriesCollection.NewSeries
    With chtChart.SeriesCollection(3 + o)
        .XValues = Xrange_O
        .Values = Yrange_O
        '.MarkerStyle = xlMarkerStyleNone
        With .Border
            .Color = RGB(200, 200, 200) '  �����
            .Weight = 1 'xlThick
        End With
    End With

Next o



'��������� �������� ����� �����
For l = 0 To 12 '6

    Set Xrange_L = Sheets(sheetname).Range(AS1.Cells(r1 + 12 + l * 2, 1), AS1.Cells(r1 + 12 + l * 2, 2))
    Set Yrange_L = Sheets(sheetname).Range(AS1.Cells(r1 + 12 + l * 2 + 1, 1), AS1.Cells(r1 + 12 + l * 2 + 1, 2))
    
    chtChart.SeriesCollection.NewSeries
    With chtChart.SeriesCollection(7 + l)
        .XValues = Xrange_L
        .Values = Yrange_L
        '.MarkerStyle = xlMarkerStyleNone
        With .Border
            .Color = RGB(200, 200, 200) '  �����
            .Weight = 1 'xlThick
        End With
    End With

Next l



'������� ��������� ������ ����� �� ������������ �����
Set chtChart = chtChart.Location(Where:=xlLocationAsObject, Name:=sheetname)

chtChart.Parent.Height = ch_H
chtChart.Parent.Width = ch_W
chtChart.Parent.Top = Range(AS1.Cells(cell_R, cell_C), AS1.Cells(cell_R, cell_C)).Top '���������� ��������
chtChart.Parent.Left = Range(AS1.Cells(cell_R, cell_C), AS1.Cells(cell_R, cell_C)).Left '���������� ��������





'��������� ������� ����������
If Frame_mode = 1 Then
    Yb1 = -Imax: Yb2 = Imax
    Xb1 = -Imax: Xb2 = Imax
End If

If Frame_mode = 2 Then
    '����� ������ ��� ������� �������
    '������������ � ����������� ��������
    Xmax = max_from_2(Application.Max(Xrange1), Application.Max(Xrange2))
    Xmin = min_from_2(Application.Min(Xrange1), Application.Min(Xrange2))
    
    Ymax = max_from_2(Application.Max(datarange1), Application.Max(datarange2))
    Ymin = min_from_2(Application.Min(datarange1), Application.Min(datarange2))
    
    
    '������� ��������� ���� ��� ������� �����������
    d_buity = 0.05
    
    
    If Abs(Ymax - Ymin) >= 2 * max_from_2(Abs(Xmin), Abs(Xmax)) Then
        Xb1 = -Abs(Ymax - Ymin) / 2: Xb2 = Abs(Ymax - Ymin) / 2
        Yb1 = Ymin: Yb2 = Ymax
    Else '!!!! ���� ���������� ������ ������� - �� �� ���� ��� ���������� �������� ������������
        max_X = max_from_2(Abs(Xmin), Abs(Xmax)): dY = Abs(Ymax - Ymin)
        Xb1 = -max_X: Xb2 = max_X
        If Abs(Ymin) >= Abs(Ymax) Then
            Yb1 = Ymin: Yb2 = Yb1 + 2 * max_X
        Else
            Yb2 = Ymax: Yb1 = Yb2 - 2 * max_X
        End If
        'Yb1 = (Ymin / dY) * 2 * max_X: Yb2 = (Ymax / dY) * 2 * max_X:
        'Yb1 = -max_X: Yb2 = max_X
    End If
    
    'Yb2 = Yb2 * (1 + d_buity): Yb1 = Yb1 * (1 + d_buity)
    'Xb2 = Xb2 * (1 + d_buity): Xb1 = Xb1 * (1 + d_buity)
    
End If

With chtChart
    .ChartType = xlXYScatterLinesNoMarkers
    .Axes(xlValue, xlPrimary).MaximumScale = Yb2
    .Axes(xlValue, xlPrimary).MinimumScale = Yb1
    .Axes(xlCategory, xlPrimary).MaximumScale = Xb2
    .Axes(xlCategory, xlPrimary).MinimumScale = Xb1
    .HasAxis(xlCategory) = False
    .HasAxis(xlValue) = False
    .HasLegend = False
    .Axes(xlValue).MajorGridlines.Delete
    .PlotArea.Interior.Color = RGB(245, 245, 245) '��� �����
End With

'��������� ������� �������� ���
LI_shape = chtChart.Shapes.Count '������ ��������� ����� �� �������
h0 = chtChart.PlotArea.InsideTop '������ ������� ���������� �� ������ ���� ���������
For o = 0 To 3
    left_shift_text = ch_W * 0.5
    top_shift_text = h0 * 0 + chtChart.PlotArea.Height * (Abs(Yb2) + Imax * (o + 1) / 4) / (Abs(Yb1) + Abs(Yb2))
    chtChart.Shapes.AddTextbox(msoTextOrientationHorizontal, left_shift_text, top_shift_text, 200, 20).TextFrame.Characters.Text = Val(Imax * (1 + o) / 4)
    With chtChart.Shapes(LI_shape + o + 1).DrawingObject.Font
        .Name = "Arial"
        .Size = 12
    End With
Next o

'����������� ��������� � �������� '
chtChart.CopyPicture
ActiveSheet.Paste Destination:=ActiveSheet.Range(AS1.Cells(cell_R, cell_C), AS1.Cells(cell_R, cell_C))


'�������� ������� (��� ������ 2003) ��� ���� ��������� (��� ������ 2007)
If Val(Application.Version) > 11 Then
    last_pasted_ind = -1
    For s = 1 To AS1.Shapes.Count
        If AS1.Shapes(s).Type = msoPicture Then last_pasted_ind = s
    Next s
    AS1.Shapes(last_pasted_ind).Name = ch_name
    ind = chtChart.Parent.Index
    AS1.ChartObjects(ind).Delete
Else
    For Each s In chtChart.SeriesCollection
        s.Delete
    Next s
    For Each s In chtChart.Shapes
       If s.Type = msoTextBox Then s.Delete
    Next s
    chtChart.Parent.Name = ch_name
End If

'�������� ������ ��� ����������
Range(AS1.Cells(r1, 1), AS1.Cells(r1 + N_r - 1, 2 * N_c - 1)).Clear

End Sub

Sub Fill_iesFile_Output(Top_Corner_Row, Top_Corner_Column, G, N_c, Ic0, Ic90)

'���������� ������� ������ "�����" � ���������� ��� - ���� ������ �� ������������� ����� �0 � �90
'ReDim G(18), Ic0(18), Ic90(18)
Ng = UBound(G)

m = Create_IES_Matrix_byEllipse090(G, N_c, Ic0, Ic90)

For i = 0 To N_c - 1
ActiveSheet.Cells(Top_Corner_Row + i + 2, Top_Corner_Column) = i * 90 / (N_c - 1) '������� ����� C - ��������������, ��� ��������� ���� �=0, � �������� - 90
    For j = 0 To Ng
        ActiveSheet.Cells(Top_Corner_Row + i + 2, Top_Corner_Column + j + 1) = m(i, j)
        If i = 0 Then ActiveSheet.Cells(Top_Corner_Row + i, Top_Corner_Column + j + 1) = G(j) '������� ����� "�����"
        If i = 1 And j = 0 Then
            For jjj = 0 To N_c - 1
                ActiveSheet.Cells(Top_Corner_Row + i, Top_Corner_Column + jjj + 1) = jjj * 90 / (N_c - 1) '������� ����� "C" ��� ies-��������
            Next jjj
        End If
    Next j
Next i

End Sub

Sub create_IES_file(FullPath, start_Row, start_col, two_waste_Rows)
'Path = "D:\2016 kuraev-pc-blt"
'Filename = "patlacco"
'FullPath = Path & "\" & Filename & ".ies"

Open FullPath For Output As #1


's = ����1.Cells(i + 1, 1)

'������������ ������
Ni = 0
'start_Row = 1: start_col = 1
i = 0: j = 0
cells_string = ""

ReDim string_output(0): string_output(0) = ""
Do While IsEmpty(ActiveSheet.Cells(i + start_Row, start_col)) = False And i < 150
    Do While IsEmpty(ActiveSheet.Cells(i + start_Row, j + start_col)) = False And j < 150
        If j = 0 Then
            cells_string = ActiveSheet.Cells(i + start_Row, j + start_col)
        Else
            cells_string = cells_string & " " & ActiveSheet.Cells(i + start_Row, j + start_col)
        End If
        j = j + 1
    Loop
    
    ReDim Preserve string_output(i): string_output(i) = Replace(cells_string, ",", ".")
    
    cells_string = ""
    i = i + 1: j = 0
Loop
Ni = i - 1 '-2 - ��� ���������� ��� ������ ����



'��������� �����
For i = 0 To Ni - 1
    Print #1, string_output(i)
Next i
Print #1, string_output(Ni); ' ";" � ����� �������� ������� �������

Close #1
End Sub

Sub take_and_save_picture(import_pic_path, diagramm_number, pic_max_height, pic_max_width, export_pic_path, pic_ext)
    Set Picture_buffer = ActiveSheet.Pictures.Insert(import_pic_path): pH = Picture_buffer.ShapeRange.Height: pW = Picture_buffer.ShapeRange.Width: Picture_buffer.Delete '��������� ������� �����������
    ActiveSheet.ChartObjects(diagramm_number).Activate '���������� ���������
    '�������� ������:
    If pic_max_width >= pic_max_height Then
        ActiveChart.Parent.Width = pic_max_width: ActiveChart.Parent.Height = (pic_max_width * pH) \ pW  '"\"-������������� �������
    Else
        ActiveChart.Parent.Height = pic_max_height: ActiveChart.Parent.Weight = (pic_max_height * pW) \ pH
    End If
    ActiveChart.ChartArea.Fill.UserPicture import_pic_path '��������� �������� (�����, �������� � ������� ��� ��������)
    'ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)
    ActiveChart.ChartArea.Border.LineStyle = 0


    ActiveChart.Export Filename:=export_pic_path, filtername:=pic_ext '��������� ��������

End Sub

Function Array_WithOut_El(ByVal Arr, ByVal el_ind) As Variant() '�������, ������� ������� ������� �� �������
ReDim A_buff(UBound(Arr) - 1)
For i = 0 To UBound(A_buff)
    
        If i < el_ind Then A_buff(i) = Arr(i)
        If i >= el_ind Then A_buff(i) = Arr(i)
    
Next i
Array_WithOut_El = A_buff
End Function



Function arythm_interp(x0, A, B, fA, fB)

    tgA = (fB - fA) / (B - A)
    arythm_interp = fA + tgA * (x0 - A)

End Function

Function ellipse_approximation(I_0, I_N, b0, c_bi, bN)
'������� ���������
'I_0, I_N - ��������� � �������� ���� �����
'b0, c_bi, bN - ��������������,

Nb = bN - b0 + 1 'eie-ai oaeia, a eioi?ii iu aii?ieneie?oai
d_b = (bN - b0) / (Nb - 1)     'oaa oaeia
i_b = c_bi

If I_0 <= I_N Then

    betta = (bN - i_b) * d_b
    betta = betta * 3.14159 / 180
    ee = Sqr(1 - (I_0 / I_N) ^ 2)
    bb = I_0
    
End If
If I_0 > I_N Then

    betta = (i_b - b0) * d_b
    betta = betta * 3.14159 / 180
    ee = Sqr(1 - (I_N / I_0) ^ 2)
    bb = I_N
    
End If

Iint = bb / Sqr(1 - (ee * Cos(betta)) ^ 2)

ellipse_approximation = Iint

End Function


Function Create_IES_Matrix_byEllipse090(gamma_ang, Nc, I0, I90) As Variant()
'!!!!! ��������� ������� ������� ��������� ������ 1000 ��!!!!!!!!!!!!!!!!!!
'������� ������ IES0 � 90 , ���� ����� � ��������, ���-�� ����� �
ReDim M_buff(UBound(gamma_ang), Nc - 1)
N_gamma = UBound(gamma_ang)
dC = (90 / (Nc - 1)) * 3.14159 / 180
F1 = 0: F2 = I0(0) * 2 * 3.14159 * (Cos(0) - Cos(gamma_ang(0) * 3.14159 / 180))
For i = 0 To N_gamma
    '������� 1 - ������ ���������� ������ �� ������� ����
    '��� �� ��������, ��� ��� ������� �� ����� ���� ����� � ���� �������,
    '�.�. ����, ��������, �� 0���� �� 5���� - ����� ���� ����� I(0;0), � ���� �� ����� I(2.5;2.5)
    '������� 2 - ������ ���������� ������, ��������� ����� ���� , ��� ���� ����� ����� ����������
    If i <> N_gamma Then
            dCos1 = Abs(Cos(gamma_ang(i) * 3.14159 / 180) - Cos(gamma_ang(i + 1) * 3.14159 / 180)) ' 1�� �������
            If i > 0 Then '2�� �������
                g1 = gamma_ang(i) - 0.5 * (gamma_ang(i - 1) + gamma_ang(i))
                g2 = gamma_ang(i) + 0.5 * (gamma_ang(i + 1) + gamma_ang(i))
                dCos2 = Abs(Cos(g1 * 3.14159 / 180) - Cos(g2 * 3.14159 / 180))
            End If
    Else
            dCos1 = Cos(gamma_ang(i - 1) * 3.14159 / 180) - Cos(gamma_ang(i) * 3.14159 / 180) ' 1�� �������
            g1 = gamma_ang(i) - 0.5 * (gamma_ang(i - 1) + gamma_ang(i)): dCos2 = Cos(g1 * 3.14159 / 180) * 2 ' 2�� �������
            
    End If
    
    
 
    
    For j = 0 To Nc - 1
        If i > 0 And i < Nc - 1 Then Ic = ellipse_approximation(I0(i), I90(i), 0, dC * j * 180 / 3.14159, 90)
        If i = 0 Then Ic = I0(i)
        If i = Nc - 1 Then Ic = I90(i)
            
        '��������� �����
        dF1 = dCos1 * dC * Ic ' 1�� �������
        dF2 = dCos2 * dC * Ic ' 2�� �������
        '������ �����
        F1 = F1 + dF1 ' 1�� �������
        F2 = F2 + dF2 ' 2�� �������
        
        
        M_buff(j, i) = Ic '���� �����
        
        
        
    Next j
Next i

F = 4 * F1 ' ������ �� ������� ��������
'F = F2 ' ������ �� ������� ��������

K1000 = 1000 / F
'K1000 = K1000 * K_DIALUX '����������� ��������, ����� �������� �������� ����� � ���, � ������������ �� ������� ���

For i = 0 To N_gamma
    For j = 0 To Nc - 1
        M_buff(i, j) = Round(M_buff(i, j) * K1000, 0)
    Next j
Next i


Create_IES_Matrix_byEllipse090 = M_buff

End Function


Function find_empty_area(start_Row, N_row, N_col) '���������� ����� ����, � �������� ���������� ��������� ������� ����� ��������� ���������

fin_row = start_Row - 1

r1 = start_Row - 1: R2 = r1 + N_row - 1: C = N_col
range_is_looking_for = True: extra_count = 1
Do While range_is_looking_for And extra_count < 1000

r1 = r1 + 1: R2 = r1 + N_row - 1
Set oRange = Range(ActiveSheet.Cells(r1, 1), ActiveSheet.Cells(R2, C))
If oRange.Text = "" Then range_is_looking_for = False

fin_row = fin_row + 1
extra_count = extra_count + 1
Loop

find_empty_area = fin_row


End Function
Function GetFilePath(Optional ByVal Title As String = "�������� ����", _
                     Optional ByVal InitialPath As String = "D:\����������") As String
On Error Resume Next
    With Application.FileDialog(msoFileDialogOpen)
        .ButtonName = "�������": .Title = Title: .InitialFileName = InitialPath
        If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1): PS = Application.PathSeparator
    End With
End Function

Sub Copy_File(sFileName, sNewFileName)
    
    'sFileName = "C:\WWW.xls"    '��� ����� ��� �����������
    'sNewFileName = "D:\WWW.xls"    '��� ����������� �����. ����������(� ������ ������ ���� D) ������ ������������
    If Dir(sFileName, 16) = "" Then MsgBox "��� ������ �����", vbCritical, "������": Exit Sub
    
    FileCopy sFileName, sNewFileName '�������� ����
    MsgBox "���� ����������", vbInformation, "www.excel-vba.ru"
End Sub



Function Convert_ChartName_to_ChartNumber(chart_name) As Integer

For s = 1 To ActiveSheet.ChartObjects.Count
    If ActiveSheet.ChartObjects(s).Name = chart_name Then Convert_ChartName_to_ChartNumber = s
Next s

End Function

Function Convert_ShapeName_to_ShapeNumber(shape_name) As Integer

For s = 1 To ActiveSheet.Shapes.Count
    If ActiveSheet.Shapes(s).Name = shape_name Then Convert_ShapeName_to_ShapeNumber = s
Next s

End Function

Function min_from_2(n1, n2)
If n1 <= n2 Then
    min_from_2 = n1
Else
    min_from_2 = n2
End If
End Function
Function max_from_2(n1, n2)
If n1 >= n2 Then
    max_from_2 = n1
Else
    max_from_2 = n2
End If
End Function
