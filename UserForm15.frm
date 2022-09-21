VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "UserForm15.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'=====================================================
'Copyright (c) Mechanic Team - RD Filter VACE. All rights reserved.
''' <Summary> '''
' Handle Data from AIB Test
' Revision 1  :


' Revision 2  : Update enter the name file and column

' Revision 3 : Update History Page


'======================================================



Private Sub btn_ChayDi_Click()
        ' *************************************************************
        ' *********** CREAT GLOBAL VARIANT  ***************************
        ' *************************************************************
            Dim FiletoOpen As Variant
            Dim demA As Integer
            Dim i As Integer
            Dim Name_file_now As String
       ' *************************************************************
       
    ' =========== get file name now ========
    Name_file_now = ThisWorkbook.Name
    
    ' =========== Enter File Excel to ======
    FiletoOpen = Application.GetOpenFilename(MultiSelect:=True) ' Luu du lieu nhung file excel nap vao
    demA = UBound(FiletoOpen) ' Lay so luong file nap
    
    
    ' ======= 1. Vong for cho lap chay nhieu doi tuong ======
    Dim FileImport As Workbook
    Dim name_file_import_i As String
    
    
    ' ======= 1.2 Import Data from Form - Sheetname, Beginning Column, Ending Column ...
        Dim Name_sheet_1 As String
        Dim Name_sheet_2 As String
        Dim Name_sheet_3 As String
        Dim Begin_Col_Sheet1 As Integer
        Dim Begin_Col_Sheet2 As Integer
        Dim End_Col_Sheet1 As Integer
        Dim End_Col_Sheet2 As Integer
        Dim Status_Export_Sheet_A As Boolean
        Dim Status_Export_Sheet_B As Boolean
        Dim Status_Export_Sheet_History As Boolean
        
        Name_sheet_1 = Tbx_Name_sheet1.Value
        Name_sheet_2 = Tbx_Name_sheet2.Value
        Name_sheet_3 = Tbx_Name_sheet3.Value
        
        Begin_Col_Sheet1 = CInt(tbx_Begin_Col_sheetA.Value)
        Begin_Col_Sheet2 = CInt(tbx_Begin_Col_sheetB.Value)
        End_Col_Sheet1 = CInt(tbx_End_Col_sheetA.Value)
        End_Col_Sheet2 = CInt(tbx_End_Col_sheetB.Value)
        Status_Export_Sheet_A = Cbx_Choose_sheetA.Value
        Status_Export_Sheet_B = Cbx_Choose_sheetB.Value
        Status_Export_Sheet_History = Cbx_Choose_sheetC.Value
        
        ' *************************************************************
        ' *********** CREAT GLOBAL VARIANT  ***************************
        ' *************************************************************
            Dim dong_cuoi As Integer
        ' *************************************************************
        
        
        
        ' *************************************************************
        ' *********** CHECK STATUS_EXPORT_SHEET_A is OK ***************
        ' *************************************************************
        If Status_Export_Sheet_A = True Then
            ' =========================================================
            ' ======= 2. Copy sheet AIB test vao sheet hien hanh ======
            ' =========================================================
            For i = 1 To demA
                Workbooks.Open FiletoOpen(i)
                Set FileImport = ActiveWorkbook
                name_file_import_i = FileImport.Name
                Sheets(Name_sheet_1).Copy before:=Workbooks(Name_file_now).Sheets(1)
                FileImport.Close False
                Workbooks(Name_file_now).Activate
                Sheets(1).Name = name_file_import_i
            Next i
        '
                ' ==========================================================
                ' ========== 3. Copy du lieu  sang sheet_1 Summary =========
                ' ==========================================================
                '
                Dim num_col As Integer
                num_col = End_Col_Sheet1 - Begin_Col_Sheet1 + 1
                'MsgBox ("So luong cot la : " & num_col)
                Dim sheet_AIB As Worksheet
                '
                Workbooks(Name_file_now).Activate
                Set sheet_AIB = Worksheets.Add
                sheet_AIB.Name = Name_sheet_1 & "Summary"
                '
                For i = 2 To demA + 1
                    Sheets(i).Select
                    ActiveSheet.Cells.UnMerge
                    dong_cuoi = Sheets(i).UsedRange.Rows(Sheets(i).UsedRange.Rows.Count).Row
                    'MsgBox ("Dong cuoi la : " & dong_cuoi)
                    Range(Cells(1, Begin_Col_Sheet1), Cells(dong_cuoi, End_Col_Sheet1)).Copy Destination:=sheet_AIB.Cells(2, num_col * i)
                    sheet_AIB.Cells(1, num_col * i) = Sheets(i).Name
                Next i
                ' == Delete sheet ==
                For i = 2 To demA + 1
                    Sheets(2).Delete
                Next i
         End If ' Finish endif "Check status export sheet A is OK "
         
         
         
         
         
        ' *************************************************************
        ' *********** CHECK STATUS_EXPORT_SHEET_B is OK ***************
        ' *************************************************************
        If Status_Export_Sheet_B = True Then
                ' ==================================================================
                ' ======= 4. Copy sheet "Sheet_2 Summary" vao sheet hien hanh ======
                ' ==================================================================
                For i = 1 To demA
                    Workbooks.Open FiletoOpen(i)
                    Set FileImport = ActiveWorkbook
                    name_file_import_i = FileImport.Name
                    Sheets(Name_sheet_2).Copy before:=Workbooks(Name_file_now).Sheets(1)
                    FileImport.Close False
                    Workbooks(Name_file_now).Activate
                    Sheets(1).Name = name_file_import_i
                Next i
        
            ' ==========================================================================
            ' ========== 5. Copy du lieu "AIB test" sang sheet_Isolation Summary =======
            ' ==========================================================================
            Dim num_col_2 As Integer
            num_col_2 = End_Col_Sheet2 - Begin_Col_Sheet2 + 1
            'MsgBox ("So luong cot la : " & num_col)
            Dim sheet_No_2 As Worksheet
            Workbooks(Name_file_now).Activate
            Set sheet_No_2 = Worksheets.Add
            sheet_No_2.Name = Name_sheet_2 & "Summary"
            For i = 2 To demA + 1
                Sheets(i).Select
                ActiveSheet.Cells.UnMerge
                dong_cuoi = Sheets(i).UsedRange.Rows(Sheets(i).UsedRange.Rows.Count).Row
                'MsgBox ("Dong cuoi la : " & dong_cuoi)
                Range(Cells(1, Begin_Col_Sheet2), Cells(dong_cuoi, End_Col_Sheet2)).Copy Destination:=sheet_No_2.Cells(2, num_col_2 * i)
                sheet_No_2.Cells(1, num_col_2 * i) = Sheets(i).Name
            Next i
            For i = 2 To demA + 1
                Sheets(2).Delete
            Next i
         End If ' Finish " Check status export sheet B is OK "
         
         
         
         
         
        ' *************************************************************
        ' *********** CHECK STATUS_EXPORT_SHEET_History is OK *********
        ' *************************************************************
     If Status_Export_Sheet_History = True Then
                ' ==================================================================
                ' ======= 6. Copy sheet "Sheet_Hist.Page" vao sheet hien hanh ======
                ' ==================================================================
                For i = 1 To demA
                    Workbooks.Open FiletoOpen(i)
                    Set FileImport = ActiveWorkbook
                    name_file_import_i = FileImport.Name
                    Sheets(Name_sheet_3).Copy before:=Workbooks(Name_file_now).Sheets(1)
                    FileImport.Close False
                    Workbooks(Name_file_now).Activate
                    Sheets(1).Name = name_file_import_i
                Next i
            ' ==========================================================================
            ' ========== 7. Copy du lieu "History Page" sang sheet_History Page Summary =======
            ' ==========================================================================
            Dim num_col_3 As Integer
            num_col_3 = 5
            'MsgBox ("So luong cot la : " & num_col)
            Dim sheet_No_3 As Worksheet
            Workbooks(Name_file_now).Activate
            Set sheet_No_3 = Worksheets.Add
            sheet_No_3.Name = Name_sheet_3 & " Summary"
            For i = 2 To demA + 1
                Sheets(i).Select
                ActiveSheet.Cells.UnMerge
                dong_cuoi = Sheets(i).UsedRange.Rows(Sheets(i).UsedRange.Rows.Count).Row
                'MsgBox ("Dong cuoi la : " & dong_cuoi)
                Range(Cells(1, 2), Cells(dong_cuoi, 6)).Copy Destination:=sheet_No_3.Cells(2, num_col_3 * i)
                sheet_No_3.Cells(1, num_col_3 * i) = Sheets(i).Name
            Next i
            For i = 2 To demA + 1
                Sheets(2).Delete
            Next i
     End If ' Finish "Check status export history is Ok
        
        
     MsgBox ("Chay xong roi nhe !!!")
        

    Unload Me
    

End Sub

Private Sub Label7_Click()

End Sub
