
Sub main()

    Dim ORDER_SHEET_NAME_ARRAY As Variant
    ORDER_SHEET_NAME_ARRAY = Array("A00", "R00", "J00")

    If (Application.Workbooks.Count >= 2) Then
        MsgBox ("不要なブックが開かれています。" & Chr(13) & _
        "このブック以外を閉じて再度実行してください。")
        Exit Sub
    End If



    MsgBox ("抽出処理を実行します。受注ファイルを選択して下さい。")
    
    
    
    Dim target_filename As String
    target_filename = get_target_filename()
    
    If (target_filename = "") Then
        MsgBox ("受注ファイルが選択されませんでした。処理を中断します。")
        Exit Sub
    End If
    
    
    
    Dim target_book As Workbook
    Set target_book = Application.Workbooks(target_filename)
    
    If (is_appropriate_target(target_book) <> True) Then
        MsgBox ("選択されたファイルは適切な受注ファイルではありません。" & _
        Chr(13) & "処理を中断します｡ ")
        
        Exit Sub
    End If
    
    
    
    
    ' probably make a sheet name check function
    
    finded_sheet_name = get_sheet_name(target_book, ORDER_SHEET_NAME_ARRAY)
    
    If (finded_sheet_name = "") Then
        MsgBox ("登録されたシート名が見つかりませんでした。" & _
        Chr(13) & "処理を中断します。")
        Exit Sub
        
    End If
    
    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(finded_sheet_name)
    
 
    

    If (is_there_rawdata_sheet(target_book) <> True) Then
        Set raw_data_sheet = target_book.Worksheets.Add(after:=Worksheets(Worksheets.Count))
        
        Call initialize_raw_data_sheet(raw_data_sheet)
        
        Dim exclusion_list_sheet As Worksheet
        Set exclusion_list_sheet = Application.ThisWorkbook.Sheets("ExclusionList")
        Call create_raw_data(raw_data_sheet, target_sheet, exclusion_list_sheet)
    Else
        MsgBox ("すでに抽出された受注データがあります。")
        Set raw_data_sheet = target_book.Worksheets("RawData")
    End If
    
    
    
    
'--------------------------------------------------
    
    
    MsgBox ("在庫管理表に受注を入力します。" & Chr(13) & "在庫管理表ファイルを詮索してください。")
    Dim management_filename As String
    management_filename = get_target_filename()
    
    If (management_filename = "") Then
        MsgBox ("在庫管理表ファイルが選択されませんでした。処理を中断します。")
        Exit Sub
    End If
    
    
    Dim management_book As Workbook
    Set management_book = Application.Workbooks(management_filename)
    
    If (is_appropriate_management_book(management_book) <> True) Then
        MsgBox ("選択されたファイルは適切な在庫管理表ファイルではありません。" & _
        Chr(13) & "処理を中断します｡ ")
        
        Exit Sub
    End If
    
    Dim management_sheet As Worksheet
    Set management_sheet = management_book.Worksheets("在庫管理表")
    
    Call input_to_management_sheet(raw_data_sheet, management_sheet)

    
'--------------------------------------------------


    MsgBox ("処理が完了しました。")
    'target_book.Close savechanges:=False
    'management_book.Close savechanges:=False
    
End Sub



Function get_target_filename()

    Dim full_target_filename As String
    full_target_filename = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
    
    
    If full_target_filename <> "False" Then
        Workbooks.Open full_target_filename
        get_target_filename = Dir(full_target_filename)
    
    Else
        get_target_filename = ""
        
    End If

End Function


Function is_appropriate_target(target_book) As Boolean


    If ((target_book.Worksheets(1).Cells(5, 6) = "品番") _
    And (target_book.Worksheets(1).Cells(2, 2).Value = "受注データ一覧")) Then
        is_appropriate_target = True
    Else
        is_appropriate_target = False
    End If
    
End Function

Function get_sheet_name(target_book, ORDER_SHEET_NAME_ARRAY) As String

    For Each name_element In ORDER_SHEET_NAME_ARRAY
        For Each a_sheet In target_book.Worksheets
            If (name_element = a_sheet.Name) Then
                get_sheet_name = a_sheet.Name
                Exit Function
            End If
        Next a_sheet
    Next name_element
        
    get_sheet_name = ""
    
End Function

Function is_there_rawdata_sheet(target_book) As Boolean

    Dim ws As Worksheet, raw_flag As Boolean

    For Each ws In target_book.Worksheets
        If ws.Name = "RawData" Then
            is_there_rawdata_sheet = True
            Exit Function
        End If
    Next ws

    is_there_rawdata_sheet = False

End Function

Function initialize_raw_data_sheet(raw_data_sheet)

    raw_data_sheet.Name = "RawData"
    raw_data_sheet.Cells(1, 1).Value = "品番"
    raw_data_sheet.Cells(1, 2).Value = "品名"
    raw_data_sheet.Cells(1, 3).Value = "数量"
    raw_data_sheet.Cells(1, 4).Value = "納期"
    raw_data_sheet.Cells(1, 5).Value = "flag"

End Function




Function get_item_count(obj_sheet, row, column, step) As Integer

    Dim item_count As Integer
    item_count = 0
    
        
    Do While (obj_sheet.Cells(row + item_count, column) <> "")
        item_count = item_count + step
    Loop
    
    get_item_count = item_count / step

End Function


Function create_raw_data(raw_data_sheet, target_sheet, exclusion_list_sheet)

    Dim target_row As Integer, raw_data_row As Integer
    
    Dim target_item_max_count As Integer, exclusion_item_max_count As Integer
    target_item_max_count = get_item_count(target_sheet, 6, 6, 1)
    exclusion_item_max_count = get_item_count(exclusion_list_sheet, 2, 1, 1)
    'MsgBox (target_item_max_count)
    
    
    
    Dim judge_item As String, exclusion_item As String
    Dim exclusion_flag As Boolean
    Dim raw_data_count As Integer
    raw_data_count = 0
    
    For i = 1 To target_item_max_count
        
        If (target_sheet.Cells(5 + i, 13).Value <> "") Then
            judge_item = target_sheet.Cells(5 + i, 6).Value
            exclusion_flag = False
                For j = 1 To exclusion_item_max_count
                    exclusion_item = exclusion_list_sheet.Cells(1 + j, 1).Value
                    
                    If (exclusion_item = judge_item) Then
                        exclusion_flag = True
                    End If
            
                Next j
            If (exclusion_flag <> True) Then
                raw_data_sheet.Cells(2 + raw_data_count, 1).NumberFormatLocal = "@"
                raw_data_sheet.Cells(2 + raw_data_count, 1).Value = judge_item
                
                raw_data_sheet.Cells(2 + raw_data_count, 2).NumberFormatLocal = "@"
                raw_data_sheet.Cells(2 + raw_data_count, 2).Value = target_sheet.Cells(5 + i, 9).Value
                
                raw_data_sheet.Cells(2 + raw_data_count, 3).NumberFormatLocal = "###,###"
                raw_data_sheet.Cells(2 + raw_data_count, 3).Value = target_sheet.Cells(5 + i, 13).Value
                
                raw_data_sheet.Cells(2 + raw_data_count, 4).NumberFormatLocal = "mm/dd"
                raw_data_sheet.Cells(2 + raw_data_count, 4).Value = target_sheet.Cells(5 + i, 18).Value
                
                raw_data_sheet.Cells(2 + raw_data_count, 5).Value = False
                raw_data_count = raw_data_count + 1
            End If
        End If
    Next i
    
    MsgBox (raw_data_count & "件の受注を抽出しました。")

End Function


Function is_appropriate_management_book(management_book) As Boolean


    If (management_book.Worksheets("在庫管理表").Cells(1, 2) = "在庫管理表") Then
        is_appropriate_management_book = True
    Else
        is_appropriate_management_book = False
    End If
    
End Function




Function input_to_management_sheet(raw_data_sheet, management_sheet)

    Dim raw_data_max_count As Integer, manage_data_max_count As Integer
    raw_data_max_count = get_item_count(raw_data_sheet, 2, 1, 1)
    manage_data_max_count = get_item_count(management_sheet, 2, 1, 3)

    
    Dim manage_month As Integer
    manage_month = Month(management_sheet.Cells(1, 1))
    
    Dim add_column As Date
    Dim item_exist_flag As Boolean
    
    Dim input_count As Integer, no_list_count As Integer, out_of_month_count As Integer
    input_count = 0
    no_list_count = 0
    out_of_month_count = 0
    
    For i = 1 To raw_data_max_count
        item_exist_flag = False
        If (raw_data_sheet.Cells(1 + i, 5).Value = False) Then
            For j = 1 To manage_data_max_count
                If (raw_data_sheet.Cells(1 + i, 1).Value = management_sheet.Cells((-1) + 3 * j, 1).Value) Then
                    If (manage_month = Month(raw_data_sheet.Cells(1 + i, 4).Value)) Then
                        add_column = raw_data_sheet.Cells(1 + i, 4) - management_sheet.Cells(1, 1).Value
                        management_sheet.Cells((-1) + j * 3, 6 + add_column) = management_sheet.Cells((-1) + j * 3, 6 + add_column) + raw_data_sheet.Cells(1 + i, 3).Value
                        For k = 1 To 5
                            raw_data_sheet.Cells(1 + i, k).Interior.Color = 16777164
                        Next k
                        item_exist_flag = True
                        raw_data_sheet.Cells(1 + i, 5).Value = True
                        input_count = input_count + 1
                    ElseIf (manage_month < Month(raw_data_sheet.Cells(1 + i, 4).Value)) Then
                        For k = 1 To 5
                            raw_data_sheet.Cells(1 + i, k).Interior.Color = 13434879
                        Next k
                        item_exist_flag = True
                        out_of_month_count = out_of_month_count + 1
                    ElseIf (manage_month > Month(raw_data_sheet.Cells(1 + i, 4).Value)) Then
                        For k = 1 To 5
                            raw_data_sheet.Cells(1 + i, k).Interior.ThemeColor = xlThemeColorDark1
                            raw_data_sheet.Cells(1 + i, k).Interior.TintAndShade = -0.149998474074526
                        Next k
                        item_exist_flag = True
                        out_of_month_count = pit_of_month_count + 1
                    End If
                
                End If

            Next j
        Else
            item_exist_flag = True
        End If
            
        If (item_exist_flag = False) Then
            For k = 1 To 5
                raw_data_sheet.Cells(1 + i, k).Interior.Color = 16764159
            Next k
            no_list_count = no_list_count + 1
        End If
    Next i

    MsgBox ("入力：" & input_count & "件, 未入力：" & no_list_count + out_of_month_count & _
    "件" & Chr(13) & "(未管理：" & no_list_count & "件, " & "管理月外：" & out_of_month_count & _
    "件)")

End Function
